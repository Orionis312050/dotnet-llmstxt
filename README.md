# dotnet-llmstxt

Édition **offline et auto-contenue** du [llms_csharp_full.txt du Teams SDK .NET](https://microsoft.github.io/teams-sdk/llms_docs/llms_csharp_full.txt), pensée pour les environnements fermés (knowledge base d'IA ne pouvant pas fetcher de liens externes).

## Contenu

- `llms.txt` (~80 Ko) : reprend le contenu complet de `llms_csharp_full.txt` (21 pages de la doc Teams SDK C#). **Chaque lien interne vers une autre page de la doc a été remplacé par une ancre in-file (`#section`)** pointant vers la section correspondante du fichier — aucun fetch externe requis.
- Les liens vraiment externes (learn.microsoft.com, github.com, images, références CLI) sont conservés tels quels.

## Source

Contenu issu du dépôt officiel [`microsoft/teams-sdk`](https://github.com/microsoft/teams-sdk) (`teams.md/docs/main`, branche `main`), traité avec la même logique que le générateur officiel `teams.md/scripts/generate-llms-txt.ts` (filtrage frontmatter `llms: ignore`, nettoyage MDX, flattening des onglets).

## Utilisation

```
https://raw.githubusercontent.com/Orionis312050/dotnet-llmstxt/main/llms.txt
```
