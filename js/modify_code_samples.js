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

    if (!fs.lstatSync(folderPath).isDirectory()) {
        return true;
    }

    //iterate each file in this folder
    var codeFiles = fs.readdirSync(folderPath);

    codeFiles.every(function (codeFile) {

        if (codeFile.indexOf("00 overview") == -1) {
            //call the function to process the string

            var codePath = path.join(folderPath, codeFile);
            var contents = cleanUpCodeFile(codePath);

            fs.writeFileSync(codePath, contents);

            console.log(contents);

        }

        return true;
    });



    return true;
}, this);

function cleanUpCodeFile(codeFile) {

    var codeFileContents = fs.readFileSync(codeFile, "utf8");

    console.log(codeFileContents);

    //add the header line to the top using filename
    var outContents = codeFileContents;
    outContents = "# code sample from " + codeFile + "\n\n"  + outContents;

    //remove the comments section

    //TODO figure out this regex

    var commentRemovalRegex = /^\s+'-{3,}[\s\S]*?-{3,}$\s+'/gm;

    outContents = outContents.replace(commentRemovalRegex, "");

    return outContents;
}