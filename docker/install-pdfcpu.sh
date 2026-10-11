#!/bin/sh
set -eu;
case "$TARGETARCH" in
  amd64) archive=Linux_x86_64; digest=757e57036e1da56b789d0243ba72b09be4f26b3ec88e781170896ac36ae04d51 ;;
  arm64) archive=Linux_arm64; digest=4719eeffbf2c8bdb099c18a7ab8dbdb9c058a546b2aa90bc6d7294be21724ece ;;
  *) exit 1 ;;
esac;
curl -fsSL "https://github.com/pdfcpu/pdfcpu/releases/download/v0.16.1/pdfcpu_0.16.1_${archive}.tar.xz" -o /tmp/pdfcpu.tar.xz;
printf '%s  %s\n' "$digest" /tmp/pdfcpu.tar.xz | sha256sum -c -;
mkdir /tmp/pdfcpu;
tar -xJf /tmp/pdfcpu.tar.xz -C /tmp/pdfcpu;
find /tmp/pdfcpu -type f -name pdfcpu -exec install -m 0755 '{}' /usr/bin/pdfcpu \;
