#!/bin/sh
# Installs Microsoft fonts into $1 when MICROSOFT_FONTS=true. Files are downloaded from Microsoft at build
# time and never stored in this repository. Their licenses allow installation on your own devices but not
# redistribution: keep an image built with this option private (see docs/site/pages/deployment.html#microsoft-fonts).
set -eu
target=$1
mkdir -p "$target"
[ "${MICROSOFT_FONTS:-false}" = true ] || exit 0

# Arial, Times New Roman, Courier New, Georgia, Verdana, Trebuchet MS, Comic Sans MS, Impact, Webdings,
# Andale Mono. The Debian contrib installer fetches Microsoft's original archives and verifies their hashes.
sed -i 's/^Components: main$/Components: main contrib/' /etc/apt/sources.list.d/debian.sources
echo 'ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true' | debconf-set-selections
apt-get update
apt-get install -y --no-install-recommends ttf-mscorefonts-installer
cp -r /usr/share/fonts/truetype/msttcorefonts "$target/core"

# Aptos family: https://www.microsoft.com/en-us/download/details.aspx?id=106087
python3 - "$target/aptos" <<'EOF'
import hashlib, io, sys, urllib.request, zipfile
url = "https://download.microsoft.com/download/8/6/0/860a94fa-7feb-44ef-ac79-c072d9113d69/Microsoft%20Aptos%20Fonts.zip"
data = urllib.request.urlopen(url, timeout=120).read()
if hashlib.sha256(data).hexdigest() != "6528fd120e719a9f985e94214eca6887d1653b88456916a792a630b02e95b025":
    raise SystemExit("Aptos download SHA-256 mismatch")
zipfile.ZipFile(io.BytesIO(data)).extractall(sys.argv[1])
EOF
