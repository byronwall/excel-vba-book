for d in ../book/*/ ; do (cd "$d" && PDF=$(echo $d | cut -d'/' -f3-) && pandoc  *.md  -V papersize:letter -o ../../builds/${PDF%?}.pdf --template=../../bin/eisvogel_chapters.latex --listings); done

