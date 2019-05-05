# excel-book

This README contains an overviews of structure and helpful scripts.

## bin scripts

Create a copy of the book.

```bash
sh create_book.sh
```

Generate the spell check dictionary

```bash
spellchecker **/*.md --generate-dictionary
```

## js scripts

Convert all of the markdown files into new sections

```bash
node rebuild-sections.js
```
