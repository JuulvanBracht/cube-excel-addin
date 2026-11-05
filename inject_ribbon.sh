#!/bin/bash
# Voegt het Cube-lint (customUI14.xml) toe aan een .xlam. Opnieuw uitvoeren mag (overschrijft).
# Gebruik:  ./inject_ribbon.sh pad/naar/Cube.xlam
set -euo pipefail

[ $# -eq 1 ] || { echo "Gebruik: $0 pad/naar/Cube.xlam"; exit 1; }
[ -f "$1" ]  || { echo "Bestand niet gevonden: $1"; exit 1; }

HIER="$(cd "$(dirname "$0")" && pwd)"
XLAM="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

cp "$XLAM" "$XLAM.bak"                       # back-up
cd "$TMP"
unzip -q "$XLAM"

mkdir -p customUI
cp "$HIER/customUI14.xml" customUI/customUI14.xml

if ! grep -q "customUI14.xml" _rels/.rels; then
  perl -0pi -e 's#</Relationships>#<Relationship Id="rIdCubeUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>#' _rels/.rels
fi

rm "$XLAM"
zip -q -X "$XLAM" '[Content_Types].xml'      # moet als eerste in de zip
zip -q -X -r "$XLAM" .
echo "Klaar: lint toegevoegd aan $XLAM (back-up: $XLAM.bak)"
