var parse = require('csv-parse/lib/sync');
var fs = require("fs");
var toMarkdown = require('to-markdown');

var input = fs.readFileSync("/Users/byronwall/Projects/excel-book/js/QueryResults (2).csv", "utf8");

var records = parse(input, {
    columns: true
});

var index = 0;

records.forEach(function (element) {


    var tags = element.tags.split("><").join(" ").split("<").join("").split(">").join("");
    console.log(element.tags, tags);

    var ans_md = toMarkdown(element.answer, {
        gfm: true
    });
    var ques_md = toMarkdown(element.question, {
        gfm: true
    });

    //output those to a fil

    var zerofilled = ('000' + index).slice(-3);
    var path = "/Users/byronwall/Projects/excel-book/book/stackover/" + zerofilled + " " + tags + ".md";

    var output = "# SO item " + zerofilled + "\n" + ques_md + "\n\n----\n\n" + ans_md + "\n";

    fs.writeFileSync(path, output);

    index++;
}, this);