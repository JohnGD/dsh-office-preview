#!/bin/sh
# Installs dsh-office-preview into a DSH / Harness home on macOS or Linux.
#
#   sh install.sh
#   DSH_HOME="$HOME/Library/Application Support/dsh-desktop/harness" sh install.sh
#
# Same three steps as install.ps1: copy the package under <home>/plugins, link it
# into <home>/profiles/node_modules, and add one `insert` row to the profile's own
# patch layer (<home>/profiles/<profile>/cordis.patch.yml).
set -eu

PLUGIN_NAME=dsh-office-preview
SOURCE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
HOME_DIR=${DSH_HOME:-}

if [ -z "$HOME_DIR" ]; then
	for candidate in \
		"$HOME/Library/Application Support/dsh-desktop/harness" \
		"$HOME/.config/dsh-desktop/harness" \
		"$HOME/.dsh"
	do
		[ -d "$candidate/profiles" ] && HOME_DIR=$candidate && break
	done
fi
[ -n "$HOME_DIR" ] && [ -d "$HOME_DIR/profiles" ] || {
	echo "Harness home not found; set DSH_HOME=<path containing profiles/>" >&2
	exit 1
}
echo "harness home : $HOME_DIR"

PROFILE=${DSH_PROFILE:-}
if [ -z "$PROFILE" ]; then
	if [ -f "$HOME_DIR/profiles/web/cordis.patch.yml" ]; then
		PROFILE=web
	else
		PROFILE=$(ls -1 "$HOME_DIR/profiles" 2>/dev/null | while read -r name; do
			[ -f "$HOME_DIR/profiles/$name/cordis.patch.yml" ] && echo "$name" && break
		done)
	fi
fi
[ -n "$PROFILE" ] || { echo "no dsh profile with a cordis.patch.yml found" >&2; exit 1; }
PATCH="$HOME_DIR/profiles/$PROFILE/cordis.patch.yml"
echo "profile      : $PROFILE"

TARGET="$HOME_DIR/plugins/$PLUGIN_NAME"
mkdir -p "$TARGET/lib"
for entry in package.json cordis.patch.yml dsh.plugin.json; do
	[ -f "$SOURCE/$entry" ] && cp -f "$SOURCE/$entry" "$TARGET/$entry"
done
for entry in "$SOURCE"/*.md; do [ -f "$entry" ] && cp -f "$entry" "$TARGET/"; done
cp -f "$SOURCE"/lib/*.js "$TARGET/lib/"
echo "files        : $TARGET"

LINK="$HOME_DIR/profiles/node_modules/$PLUGIN_NAME"
mkdir -p "$(dirname -- "$LINK")"
rm -rf "$LINK"
ln -s "$TARGET" "$LINK"
echo "link         : $LINK -> $TARGET"

if grep -q "$PLUGIN_NAME" "$PATCH"; then
	echo "patch layer  : already names $PLUGIN_NAME; left unchanged"
else
	STAMP=$(date +%Y%m%d-%H%M%S)
	cp "$PATCH" "$PATCH.bak-$STAMP"
	# Drop a bare `[]` placeholder, then append the insert row.
grep -v -x -E '[[:space:]]*\[[[:space:]]*\][[:space:]]*' "$PATCH" > "$PATCH.tmp" || true
	{
		cat "$PATCH.tmp"
		printf '\n# Office preview (Word .docx / Excel .xlsx) for the sidebar document pane.\n- insert:\n    - id: %s\n      name: %s\n' "$PLUGIN_NAME" "$PLUGIN_NAME"
	} > "$PATCH"
	rm -f "$PATCH.tmp"
	echo "patch layer  : updated $PATCH (backup: $PATCH.bak-$STAMP)"
fi

echo ''
echo "done. Refresh the Harness page. If the renderer does not appear, restart DSH."
