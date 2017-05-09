var fs = require("fs");

var path = require("path");

var srcPath = "./original code/";
var files = fs.readdirSync(srcPath);

files.every(function (element) {

    console.log(element);

    if (element === ".DS_Store") {
        return true;
    }

    var fileName = path.join(srcPath, element);

    var contents = fs.readFileSync(fileName, "utf8");

    var subRegex = /^(?:Public\s?)?(?:Sub|Function)+ (.*?)\([\s\S]*?End (Sub|Function)/gm;

    //create new directory
    var folder = path.join("./book/", element.split(".")[0]);

    if (!fs.existsSync(folder)) {
        fs.mkdirSync(folder);
    }

    var match = subRegex.exec(contents);
    while (match !== null) {

        //deal with match

        var newFile = path.join(folder, match[1] + ".md");
        var newFileContents = "```vb\n" + match[0] + "\n```";



        fs.writeFileSync(newFile, newFileContents);

        match = subRegex.exec(contents);
    }

    return true;
}, this);