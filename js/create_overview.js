var fs = require("fs");

var path = require("path");

var srcPath = "./book/";

var files = fs.readdirSync(srcPath);

files.every(function (element) {

    console.log(element);

    if (element === ".DS_Store" || element === "md code samples") {
        return true;
    }

    console.log("create overvieW");

    var folderPath = path.join(srcPath, element);

    if (!fs.lstatSync(folderPath).isDirectory) {
        return true;
    }

    //check for overview

    var overviewFile = path.join(folderPath, "overview.md");

    if (fs.existsSync(overviewFile)) {
        return true;
    }

    var overviewContents = "# overview of " + element;

    fs.writeFileSync(overviewFile, overviewContents);

    return true;
}, this);