console.log("will rename the overview files");

var fs = require("fs");
var path = require("path");

var srcPath = "./book/";

var files = fs.readdirSync(srcPath);

files.every(function (element) {

    console.log(element);

    if (element === ".DS_Store" || element === "md code samples") {
        return true;
    }

    console.log("rename overviewW");

    var folderPath = path.join(srcPath, element);

    if (!fs.lstatSync(folderPath).isDirectory) {
        return true;
    }

    //check for overview

    var overviewFile = path.join(folderPath, "overview.md");

    if (fs.existsSync(overviewFile)) {
        //do the rename
        fs.rename(overviewFile, overviewFile.replace("overview", "00 overview"));

        return true;
    }

    return true;
}, this);