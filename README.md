# dotnet-llmstxt

Édition **offline** de la doc C# (.NET) du Teams SDK, à partir de https://microsoft.github.io/teams-sdk/llms_docs/llms_csharp.txt (+ contenu complet de `llms_csharp_full.txt`).

## Structure

- `llms.txt` : index de toutes les pages (liens relatifs vers `docs/`).
- `docs/<categorie>/<feature>.md` : **une feature par fichier**, contenu complet inline.
- Chaque dossier de catégorie contient un `_category_.json` (compatible Docusaurus).
- Tous les cross-references sont des liens relatifs intra-dépôt : aucun fetch externe requis.
- `script/compilers.sh` : recompile `docs/` + l'index depuis le llms.txt complet offline.
- `script/llms.txt.part-*` : le llms.txt **complet** (liens déjà inline, ancres `#section`), poussé en parties ; `script/compilers.sh` le réassemble automatiquement en `script/llms.txt`.

## Utilisation (environnement fermé)

```bash
git clone <ce dépôt> && cd dotnet-llmstxt
bash script/compilers.sh            # réassemble script/llms.txt puis compile docs/ + index llms.txt
bash script/compilers.sh --reassemble   # réassemble uniquement script/llms.txt
bash script/compilers.sh --from-docs    # reconstruit script/llms.txt depuis docs/
```

## Source

Généré depuis `microsoft/teams-sdk` (`teams.md/src/pages/templates` + fragments `csharp.incl.md`, et `teams.md/docs/main`), même pipeline que `generate-llms-txt.ts`.
