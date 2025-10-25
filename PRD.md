# PRD - Markdown2Word Minimal

## Objectif
Convertir du texte Markdown saisi dans une interface en un fichier Word (.docx) téléchargeable.

## Public cible
Développeurs, rédacteurs ou étudiants qui veulent rapidement transformer du Markdown en Word.

## Fonctionnalités
1. **Zone de texte Markdown**
   - Saisie ou collage du texte Markdown.
   - Mise à jour en temps réel.

2. **Aperçu Markdown**
   - Affichage du rendu Markdown à droite de la zone de texte.

3. **Bouton “Convertir en Word”**
   - Génère un fichier .docx à partir du Markdown.
   - Téléchargement direct côté client.

## Workflow utilisateur
1. L’utilisateur colle ou écrit du Markdown dans la textarea gauche.  
2. L’aperçu Markdown apparaît instantanément à droite.  
3. L’utilisateur clique sur “Convertir en Word”.  
4. Le fichier Word est généré et téléchargeable.

## Technologies
- **Frontend** : Blazor WebAssembly .NET 9.0
- **Markdown Renderer** : Markdig https://www.nuget.org/packages/Markdig/0.43.0
- **Word Generator** : DocX ou Open XML SDK

## GitHub Pages
GitHub Action deploy.yml de l'application Blazor Wasm

## Critères de succès
- Conversion fidèle du Markdown de base (titres, listes, code, gras/italique).  
- Interface responsive et simple.  
- Téléchargement rapide (< 2s pour texte < 1 Mo).
