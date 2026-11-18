The Google Fonts snapshot and SHA-256 are recorded in ../sources.lock.json
(also installed under /usr/local/share/imprint/fonts in each runtime image).

Upstream groups fonts into ofl, apache, and ufl license directories:
https://github.com/google/fonts/blob/6157dfaa640ca5738b4d1a7070fe389c3036d7fc/README.md

Original font binaries, their copyright/license metadata, family license and
notice files, and METADATA.pb are preserved. These shared license texts also
cover the few families whose upstream directories omit a separate license file.
The original binaries retain the applicable font copyright notices; the shared
OFL text omits the example family's copyright heading.

Shared license text sources, all from the same pinned revision:
  OFL-1.1.txt: ofl/abeezee/OFL.txt (standard terms only)
  Apache-2.0.txt: apache/robotoslab/LICENSE.txt
  UFL-1.0.txt: ufl/ubuntu/UFL.txt

The Carbone reference list is an inventory of names, not a license grant.
Additional fonts supplied by operators keep their own licenses and notices.

Debian font packages (../debian-packages.txt) keep their licenses in
/usr/share/doc/PACKAGE/copyright inside each image.

Generated faces under /usr/local/share/fonts/imprint/generated/ are static
instances of the variable fonts above, made at build time for Carbone reference
names. Each family directory carries the source family's license files. Fonts
whose license declares a Reserved Font Name are not modified.
