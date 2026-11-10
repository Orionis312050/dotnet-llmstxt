#!/usr/bin/env bash
#
# compilers.sh — compile un llms.txt complet (offline, liens déjà inline) en arborescence docs/
#                une feature par fichier + _category_.json Docusaurus par dossier.
#
# Usage :
#   script/compilers.sh [llms.txt]   Compile le llms.txt donné (défaut: script/llms.txt) vers docs/ + index docs/llms.txt.
#   script/compilers.sh --from-docs  Reconstruit le llms.txt complet (script/llms.txt) depuis docs/ existant.
#   script/compilers.sh --reassemble  Reassemble script/llms.txt depuis script/llms.txt.part-* (si le dépôt a été poussé en parties).
#
# En environnement fermé : aucun fetch externe, tout est déjà dans le llms.txt.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DOCS_DIR="$ROOT_DIR/docs"
CAT_DIR="main-documentation"

slugify() {
  echo "$1" | tr '[:upper:]' '[:lower:]' | sed -e 's/[^a-z0-9 _-]//g' -e 's/ /-/g' -e 's/--*/-/g' -e 's/^-//' -e 's/-$//' | cut -c1-80
}

reassemble() {
  if [ ! -f "$SCRIPT_DIR/llms.txt" ] && compgen -G "$SCRIPT_DIR/llms.txt.part-*" > /dev/null; then
    cat "$SCRIPT_DIR"/llms.txt.part-* > "$SCRIPT_DIR/llms.txt"
    echo "llms.txt réassemblé depuis les parties ($(wc -c < "$SCRIPT_DIR/llms.txt") octets)."
  fi
}

compile_llms() {
  local input="$1"
  [ -f "$input" ] || { echo "llms.txt introuvable: $input" >&2; exit 1; }
  reassemble
  rm -rf "$DOCS_DIR"
  mkdir -p "$DOCS_DIR/$CAT_DIR"
  awk -v dir="$DOCS_DIR/$CAT_DIR" '
    function slug(t) {
      s = tolower(t)
      gsub(/[^a-z0-9 _-]/, "", s)
      gsub(/ /, "-", s)
      gsub(/-+/, "-", s)
      sub(/^-/, "", s); sub(/-$/, "", s)
      return substr(s, 1, 80)
    }
    /^[[:space:]]*```/ { fence = !fence; if (file != "") print > file; next }
    !fence && /^---$/ { next }
    !fence && /^### / { if (held != "" && file != "") print held > file; held = $0; next }
    {
      if (!fence && held != "" && $0 != "") {
        if ($0 ~ /^# /) {
          title = substr($0, 3); gsub(/^ +| +$/, "", title)
          if (held == "### " title) {
            if (file != "") close(file)
            base = slug(title); s = base; n = 2
            while (s in seen) { s = base "-" n; n++ }
            seen[s] = 1
            file = dir "/" s ".md"
            print "---\ntitle: \"" title "\"\n---\n" > file
            print "# " title > file
            held = ""
            next
          }
        }
        if (file != "") print held > file
        held = ""
      }
      if (file != "") print > file
    }
    END { if (file != "") close(file) }
  ' "$input"
  # _category_.json Docusaurus
  printf '{\n  "label": "Main Documentation",\n  "position": 1,\n  "collapsed": false\n}\n' > "$DOCS_DIR/$CAT_DIR/_category_.json"
  # index llms.txt dans docs/
  {
    echo "# Teams SDK - C# Documentation (offline)"
    echo ""
    echo "> Édition offline de https://microsoft.github.io/teams-sdk/llms_docs/llms_csharp_full.txt (+ contenu complet)."
    echo "> Chaque entrée pointe vers un fichier local de docs/ — aucun fetch externe requis."
    echo ""
    echo "## Main Documentation"
    echo ""
    find "$DOCS_DIR/$CAT_DIR" -maxdepth 1 -name '*.md' | sort | while read -r f; do
      t="$(sed -n 's/^title: "\(.*\)"$/\1/p' "$f" | head -1)"
      echo "- [$t]($CAT_DIR/$(basename "$f"))"
    done
  } > "$DOCS_DIR/llms.txt"
  echo "Compilation terminée: $(find "$DOCS_DIR" -name '*.md' | wc -l) fichiers dans $DOCS_DIR, index: $DOCS_DIR/llms.txt"
}

from_docs() {
  [ -d "$DOCS_DIR" ] || { echo "docs/ introuvable" >&2; exit 1; }
  out="$SCRIPT_DIR/llms.txt"
  {
    echo "# Teams SDK - C# Documentation (Complete, offline)"
    echo ""
    echo "> Complete offline edition of https://microsoft.github.io/teams-sdk/llms_docs/llms_csharp_full.txt:"
    echo "> every documentation page is inlined below as its own section. Internal links are in-file \`#anchors\`."
    echo ""
    echo "## Main Documentation"
    echo ""
    find "$DOCS_DIR/$CAT_DIR" -maxdepth 1 -name '*.md' | sort | while read -r f; do
      t="$(sed -n 's/^title: "\(.*\)"$/\1/p' "$f" | head -1)"
      echo "### $t"
      echo ""
      sed -e '1,4d' "$f"
      echo ""
      echo "---"
      echo ""
    done
  } > "$out"
  echo "llms.txt complet reconstruit: $out ($(wc -c < "$out") octets)"
}

case "${1:-}" in
  --from-docs)  from_docs ;;
  --reassemble) reassemble ;;
  "")           reassemble; compile_llms "$SCRIPT_DIR/llms.txt" ;;
  -h|--help)    sed -n '2,12p' "${BASH_SOURCE[0]}" ;;
  *)            compile_llms "$1" ;;
esac
