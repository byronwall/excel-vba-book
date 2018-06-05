const glob = require("glob");
const fs = require("fs");
const numeral = require("numeral");

function createConcatFile() {
  const fileContents = [];

  glob("../book/**/*.md", (err, files) => {
    files.forEach(file => {
      console.log(file);
      fileContents.push(fs.readFileSync(file, "utf-8"));
    });

    const result = fileContents.join("\n");
    fs.writeFileSync("../newbook/concat.md", result, "utf-8");

    // take that file and read it into lines
  });
}

function splitConcatIntoSections() {
  const allLines = fs.readFileSync("../newbook/concat.md", "utf-8");

  const lines = allLines.split("\n");

  var chapter = 0;
  var section = 0;
  var subsection = 0;

  var currentLines = [];
  var currentFile = "";
  var currentPath = "";

  // TODO: add a running array for the current file
  // TODO: determine the name of that file
  // TODO: write a new file each time a new section is hit that starts a new one

  lines.forEach(line => {
    // match line against markdown header section

    const firstHeader = /^\s*#(?!#)\s*(.*)/gm;
    const secondHeader = /^\s*##(?!#)\s*(.*)/gm;
    const thirdHeader = /^\s*###(?!#)\s*(.*)/gm;

    if (firstHeader.test(line)) {
      if (currentFile != "") {
        var textToWrite = currentLines.join("\n");
        fs.writeFileSync(currentPath + currentFile, textToWrite, "utf-8");
        currentLines = [];
      }

      chapter++;
      section = 0;
      subsection = 0;

      firstHeader.lastIndex = 0;
      var res = firstHeader.exec(line);

      var chapterName = res[1];

      var path =
        "../newbook/" +
        numeral(chapter).format("00") +
        "-" +
        cleanName(chapterName) +
        "/";

      if (!fs.existsSync(path)) {
        fs.mkdirSync(path);
      }

      currentPath = path;
      currentFile =
        numeral(chapter).format("00") + " " + cleanName(chapterName) + ".md";

      // this needs to create a new folder and set current header
    } else if (secondHeader.test(line)) {
      // write current file
      var textToWrite = currentLines.join("\n");
      fs.writeFileSync(currentPath + currentFile, textToWrite, "utf-8");
      currentLines = [];

      section++;
      subsection = 0;

      secondHeader.lastIndex = 0;
      var res = secondHeader.exec(line);

      var sectionName = res[1];

      currentFile =
        numeral(chapter).format("00") +
        "-" +
        numeral(section).format("00") +
        " " +
        cleanName(sectionName) +
        ".md";

      // write the previous file which contains # header info
    } else if (thirdHeader.test(line)) {
      subsection++;

      var textToWrite = currentLines.join("\n");
      fs.writeFileSync(currentPath + currentFile, textToWrite, "utf-8");
      currentLines = [];

      thirdHeader.lastIndex = 0;
      var res = thirdHeader.exec(line);

      var subsectionName = res[1];

      currentFile =
        numeral(chapter).format("00") +
        "-" +
        numeral(section).format("00") +
        "-" +
        numeral(subsection).format("00") +
        " " +
        cleanName(subsectionName) +
        ".md";

      // write the previous file which contains # header info
    }

    currentLines.push(line);
  });
}

var cleanName = function(name) {
  name = name.replace(/\s+/gi, "-"); // Replace white space with dash
  return name.replace(/[^a-zA-Z0-9\-]/gi, ""); // Strip any special charactere
};

splitConcatIntoSections();
