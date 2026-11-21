#!/usr/bin/env python3
"""UpdateHub release signing (maintainer tool, never installed on user machines).

  genkey         create a new signing key: passphrase-encrypted PEM + the public modulus to pin in engine.ps1
  sign           write release.signed for the files in the repo
  verify-remote  after pushing: fetch release.signed and every file exactly as clients will, and verify them

release.signed = line 1: base64 RSA-3072 PKCS#1 v1.5 / SHA-256 signature, then the manifest JSON it signs (version, tag,
size + SHA-256 of every self-updatable file). One file, so a CDN can never pair a manifest with a different signature.

    python3 -I tools/sign_release.py genkey --out ~/updatehub-signing-key.pem
    python3 -I tools/sign_release.py sign --repo . --version 1.7 --key ~/updatehub-signing-key.pem --notes "..."
    git commit ...; git push origin HEAD:refs/heads/v1.7; git push origin main      # release ref FIRST, then main
    (clients fetch files from raw.githubusercontent.com/<repo>/v1.7/..., which resolves a branch or a tag named v1.7)
    python3 -I tools/sign_release.py verify-remote --repo . --version 1.7

An encrypted key asks for its passphrase (or reads UPDATEHUB_KEY_PASSPHRASE). The key file never goes into the repository
and never leaves the maintainer's PC: sign there.
Version numbers only ever go up: clients refuse anything older than the newest signed version they have seen.
"""
import argparse, base64, getpass, hashlib, json, os, re, shutil, subprocess, sys, urllib.error, urllib.request
from datetime import date
from pathlib import Path

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa

SELF_FILES = ['engine.ps1', 'server.ps1', 'updater.html', 'README.txt', 'run.bat']   # must equal $script:SelfFiles
REQUIRED = ['engine.ps1', 'server.ps1', 'updater.html']
SLUG = 'dudidd-wq/updatehub'
BOM = b'\xef\xbb\xbf'
VER_RE = r'[0-9]{1,4}\.[0-9]{1,4}(\.[0-9]{1,4})?'


def passphrase(confirm=False):
    p = os.environ.get('UPDATEHUB_KEY_PASSPHRASE')
    if not p:
        p = getpass.getpass('Key passphrase: ')
        if confirm and p != getpass.getpass('Again: '):
            sys.exit('passphrases differ')
    if len(p) < 16:
        sys.exit('use a passphrase of at least 16 characters')
    return p.encode()


def load_key(path):
    pem = path.read_bytes()
    if b'ENCRYPTED' in pem:
        return serialization.load_pem_private_key(pem, password=passphrase())
    print('note: unencrypted key file - it must stay on this PC only (protected by your Windows account)')
    return serialization.load_pem_private_key(pem, password=None)


def b64n(n):
    return base64.b64encode(n.to_bytes((n.bit_length() + 7) // 8, 'big')).decode()


def pinned_in(engine_text):
    """(modulus b64, exponent b64) pinned in an engine.ps1 text."""
    n = re.search(r"(?m)^\$script:ReleaseKeyN\s*=\s*'([A-Za-z0-9+/=]+)'", engine_text)
    e = re.search(r"(?m)^\$script:ReleaseKeyE\s*=\s*'([A-Za-z0-9+/=]+)'", engine_text)
    if not n or not e:
        sys.exit('engine.ps1: $script:ReleaseKeyN / $script:ReleaseKeyE not found')
    return n.group(1), e.group(1)


def as_key(pin):
    n, e = (int.from_bytes(base64.b64decode(x), 'big') for x in pin)
    return rsa.RSAPublicNumbers(e, n).public_key()


def vtuple(v):
    t = tuple(int(x) for x in v.split('.'))
    return t + (0,) * (3 - len(t))


def get_url(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers={'Cache-Control': 'no-cache'}), timeout=60).read()


def published(repo):
    """What clients see now: (version, engine.ps1 of that release) or (None, None) before the first release.
    GitHub is not trusted here either: the key comes from the release's engine.ps1 in YOUR git history (tag v<ver>),
    the published release.signed must verify with it, and that engine.ps1 must match the signed hash."""
    base = f'https://raw.githubusercontent.com/{SLUG}/'
    try:
        sig, body = split_signed(get_url(base + 'main/release.signed'))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, None
        raise
    m = json.loads(body); ver = m['version']
    r = None
    for ref in (f'v{ver}', f'origin/v{ver}'):                 # local tag/branch first, then the fetched release branch
        r = subprocess.run(['git', '-C', str(repo), 'show', f'{ref}:engine.ps1'], capture_output=True)
        if not r.returncode:
            break
    if r.returncode:
        sys.exit(f'GitHub shows v{ver}, but this repo has no v{ver} (run: git fetch origin v{ver}).')
    # A forged release ref can only make this check refuse to sign (its key will not match yours) — never sign for it.
    engine_bytes = r.stdout
    as_key(pinned_in(engine_bytes.decode('utf-8-sig'))).verify(sig, body, padding.PKCS1v15(), hashes.SHA256())
    ent = [f for f in m['files'] if f['name'] == 'engine.ps1']
    if not ent or hashlib.sha256(engine_bytes).hexdigest() != ent[0]['sha256']:
        sys.exit(f'tag v{ver}:engine.ps1 does not match the published signed manifest - stop and investigate')
    return ver, engine_bytes.decode('utf-8-sig')


def check_file(name, data, version):
    """The same consistency checks engine.ps1 runs before installing (engine.ps1 also runs the PowerShell parser)."""
    text = (data[3:] if data.startswith(BOM) else data).decode('utf-8')     # strict: raises on invalid UTF-8
    v = re.escape(version)
    if name.endswith('.ps1') and not data.startswith(BOM):
        sys.exit(f'{name}: missing UTF-8 BOM (Windows PowerShell 5.1 would garble the Hebrew)')
    if name == 'engine.ps1' and not re.search(rf"(?m)^\$script:AppVersion\s*=\s*'{v}'", text):
        sys.exit(f"engine.ps1: $script:AppVersion is not '{version}'")
    if name == 'server.ps1' and not re.search(rf"(?m)^\$ServerVersion\s*=\s*'{v}'", text):
        sys.exit(f"server.ps1: $ServerVersion is not '{version}'")
    if name == 'updater.html' and (f'<meta name="updatehub-version" content="{version}">' not in text or '</html>' not in text):
        sys.exit(f'updater.html: version meta is not {version}')


def parse_check(repo):
    """Windows PowerShell 5.1's parser when available (pwsh 7 accepts syntax 5.1 rejects). Clients check again anyway."""
    pw = shutil.which('powershell.exe') or shutil.which('powershell')
    if not pw:
        if shutil.which('pwsh'):
            print('WARNING: only PowerShell 7 here; it accepts syntax (??, &&, ?:) that 5.1 rejects. Not used.')
        print('note: no Windows PowerShell here - every client parses the files with 5.1 before it installs')
        return
    for name in ('engine.ps1', 'server.ps1'):
        cmd = ("$e=$null;[void][System.Management.Automation.Language.Parser]::ParseFile('%s',[ref]$null,[ref]$e);"
               "if($e){$e|%%{$_.ToString()};exit 1}" % str((repo / name).resolve()).replace("'", "''"))
        r = subprocess.run([pw, '-NoProfile', '-NonInteractive', '-Command', cmd], capture_output=True, text=True)
        if r.returncode:
            sys.exit(f'{name}: PowerShell parse errors\n{r.stdout}{r.stderr}')


def split_signed(raw):
    nl = raw.index(b'\n')
    return base64.b64decode(raw[:nl].strip(), validate=True), raw[nl + 1:]


def cmd_genkey(a):
    out = Path(a.out).expanduser()
    if out.exists():
        sys.exit(f'{out} exists - refusing to overwrite a signing key')
    if re.search(r'onedrive|dropbox|google ?drive|icloud|box sync', str(out.resolve()), re.I):
        print('WARNING: this looks like a cloud-synced folder; keep the signing key on local disk or offline media.')
    key = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    pem = key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                            serialization.BestAvailableEncryption(passphrase(confirm=True)))
    fd = os.open(out, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(fd, 'wb') as f:
        f.write(pem)
    print(f'key written to {out} (encrypted)')
    print('pin this in engine.ps1 as $script:ReleaseKeyN:')
    print(b64n(key.public_key().public_numbers().n))


def cmd_sign(a):
    if not re.fullmatch(VER_RE, a.version):
        sys.exit('version must look like 1.7 or 1.7.1')
    repo = Path(a.repo)
    new_pin = pinned_in((repo / 'engine.ps1').read_bytes().decode('utf-8-sig'))
    # Installed copies verify with the key pinned in the release they RUN — the published one, not this working tree.
    if a.offline:
        pub_ver, pub_engine = None, None
        print('WARNING --offline: not compared with the published release (version order, pinned key).')
    else:
        pub_ver, pub_engine = published(repo)
    if pub_ver:
        if vtuple(a.version) <= vtuple(pub_ver):
            sys.exit(f'v{a.version} is not newer than the published v{pub_ver}: clients refuse it (and a re-signed '
                     'same version would be rejected as a replay). Bump the version.')
        pv, nv = vtuple(pub_ver), vtuple(a.version)
        if (nv[0] - pv[0] > 1 or (nv[0] == pv[0] and nv[1] - pv[1] > 5)) and not a.allow_jump:
            sys.exit(f'v{pub_ver} -> v{a.version} is a big jump; every client would remember it as the newest version. '
                     'Pass --allow-jump if it is intended.')
        client_pin = pinned_in(pub_engine)
    else:
        client_pin = new_pin                      # first release: nothing installed verifies against an older key
    key = load_key(Path(a.key).expanduser())
    nums = key.public_key().public_numbers()
    if (b64n(nums.n), b64n(nums.e)) != client_pin:
        sys.exit('this key is not the one installed copies trust (pinned in the published engine.ps1). They would '
                 'reject the release. Sign with that key.')
    if new_pin != client_pin:
        if not a.rotate:
            sys.exit('this release pins a NEW key in engine.ps1. Pass --rotate to confirm the rotation.')
        print('key rotation: signed with the currently trusted key; the next release must be signed with the new one.')
    files = []
    for name in SELF_FILES:
        data = (repo / name).read_bytes()
        check_file(name, data, a.version)
        files.append({'name': name, 'size': len(data), 'sha256': hashlib.sha256(data).hexdigest()})
    parse_check(repo)
    manifest = {'app': 'UpdateHub', 'version': a.version, 'tag': 'v' + a.version,
                'released': date.today().isoformat(), 'notes': a.notes[:300], 'files': files}
    body = (json.dumps(manifest, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
    sig = key.sign(body, padding.PKCS1v15(), hashes.SHA256())
    key.public_key().verify(sig, body, padding.PKCS1v15(), hashes.SHA256())
    (repo / 'release.signed').write_bytes(base64.b64encode(sig) + b'\n' + body)
    print(f'release.signed v{a.version}: {len(files)} files signed with the key installed copies trust')


def cmd_verify_remote(a):
    # Verify with the key the PREVIOUS release pinned (what installed copies use); --trusted-engine overrides,
    # and for the very first release the key in this repo's engine.ps1 is the only one there is.
    repo = Path(a.repo)
    if a.trusted_engine:
        pin = pinned_in(Path(a.trusted_engine).read_bytes().decode('utf-8-sig'))
    else:
        pin = pinned_in((repo / 'engine.ps1').read_bytes().decode('utf-8-sig'))
        print('verifying with the key pinned in this repo (pass --trusted-engine <previous engine.ps1> after a key rotation)')
    pub = as_key(pin)
    get = get_url
    base = f'https://raw.githubusercontent.com/{SLUG}/'
    sig, body = split_signed(get(base + 'main/release.signed'))
    pub.verify(sig, body, padding.PKCS1v15(), hashes.SHA256())
    m = json.loads(body)
    if m['version'] != a.version:
        sys.exit(f"main/release.signed is v{m['version']}, expected v{a.version} (CDN cache? retry in a few minutes)")
    for f in m['files']:
        data = get(base + m['tag'] + '/' + f['name'])
        if len(data) != f['size'] or hashlib.sha256(data).hexdigest() != f['sha256']:
            sys.exit(f"{f['name']} at {m['tag']} does not match the signed manifest")
        check_file(f['name'], data, m['version'])
    print(f"OK: v{m['version']} signature valid, {len(m['files'])} files at {m['tag']} match")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    g = sub.add_parser('genkey'); g.add_argument('--out', required=True)
    s = sub.add_parser('sign'); s.add_argument('--repo', required=True); s.add_argument('--version', required=True)
    s.add_argument('--key', required=True); s.add_argument('--notes', default=''); s.add_argument('--rotate', action='store_true')
    s.add_argument('--allow-jump', action='store_true'); s.add_argument('--offline', action='store_true')
    v = sub.add_parser('verify-remote'); v.add_argument('--repo', required=True); v.add_argument('--version', required=True)
    v.add_argument('--trusted-engine')
    a = ap.parse_args()
    {'genkey': cmd_genkey, 'sign': cmd_sign, 'verify-remote': cmd_verify_remote}[a.cmd](a)


if __name__ == '__main__':
    main()
