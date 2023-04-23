var assert = require("assert");
var path = require("path");
var fs = require("fs");
var _ = require("underscore");

var mammoth = require("../");
var promises = require("../lib/promises");
var results = require("../lib/results");

var testing = require("./testing");
var test = require("./test")(module);
var testData = testing.testData;
var createFakeDocxFile = testing.createFakeDocxFile;

const unzip = require('../lib/unzip')
const xmlreader = require("../lib/xml/reader");
const xml = require("../lib/xml");
const {isEmptyRun} = require("./docx/document-matchers");
var hamjest = require("hamjest");
var assertThat = hamjest.assertThat;
const {Styles} = require("../lib/docx/styles-reader");
const {defaultNumbering} = require("../lib/docx/numbering-xml");
const {createBodyReader} = require("../lib/docx/body-reader");

var allOf = hamjest.allOf;
var contains = hamjest.contains;
var hasProperties = hamjest.hasProperties;
var willBe = hamjest.willBe;
var FeatureMatcher = hamjest.FeatureMatcher;

var documentMatchers = require("./docx/document-matchers");
var isHyperlink = documentMatchers.isHyperlink;
var isRun = documentMatchers.isRun;
var isText = documentMatchers.isText;
var isTable = documentMatchers.isTable;
var isRow = documentMatchers.isRow;

var _readNumberingProperties = require("../lib/docx/body-reader")._readNumberingProperties;
var documents = require("../lib/documents");
var XmlElement = xml.Element;
var Relationships = require("../lib/docx/relationships-reader").Relationships;
var warning = require("../lib/results").warning;

const cheerio = require('cheerio')

function readXmlElement(element, options) {
    options = Object.create(options || {});
    options.styles = options.styles || new Styles({}, {});
    options.numbering = options.numbering || defaultNumbering;
    return createBodyReader(options).readXmlElement(element);
}

function readXmlElementValue(element, options) {
    var result = readXmlElement(element, options);
    assert.deepEqual(result.messages, []);
    return result.value;
}

// test('读取word并转化文件为xml', async() => {
//     const wordDocxPath = path.join(__dirname, "test-data/w_word2.docx");
//     const zip1 = await unzip.openZip({path: wordDocxPath})
//     const wordContentB = await zip1.read('word/document.xml')
//     // const wordContent = wordContentB.toString('utf-8');
//
//     // const wpsDocxPath = path.join(__dirname, "test-data/w_wps.docx");
//     // const zip2 = await unzip.openZip({path: wpsDocxPath})
//     // const wpsContentB = await zip2.read('word/document.xml')
//     // const wpsContent = wpsContentB.toString('utf-8');
//     //
//     const filepath = path.join(__dirname, "test-data/w_word2.xml");
//     fs.writeFileSync(filepath, wordContentB, {encoding: 'binary'})
//
//     // assert.equal(wordContent, wpsContent)
// })
//
test('word和wps文件转换结果相同', async function() {
    const wordDocxPath = path.join(__dirname, "test-data/w_word2.docx");
    const result1 = await mammoth.convertToHtml({path: wordDocxPath});
    const $ = cheerio.load(result1.value);
    const elements = $('FormField');
    for(let i = 0; i < elements.length; i++) {
        const element = elements[i];
        console.log(`${i+1}: ${element.attribs.id}: ${element.attribs.value}`)

    }
    //
    // const wpsDocxPath = path.join(__dirname, "test-data/w_wps.docx");
    // const result2 = await mammoth.convertToHtml({path: wpsDocxPath});

    // console.log(result1)
    // console.log(result2)
    // assert.equal(result1.value, result2.value)
});
//
// test('支持读取bookmark内的内容', async function() {
//     const bookmarkStart = await xmlreader.readString("<p><w:bookmarkStart w:id=\"0\" w:name=\"name\"/><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:fldChar w:fldCharType=\"begin\"><w:ffData><w:name w:val=\"name\"/><w:enabled/><w:calcOnExit w:val=\"0\"/><w:textInput><w:maxLength w:val=\"20\"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:instrText xml:space=\"preserve\">FORMTEXT</w:instrText></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:t>梁帅</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"default\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:t>     </w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:kern w:val=\"2\"/><w:sz w:val=\"21\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"ar-SA\"/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r><w:bookmarkEnd w:id=\"0\"/></p>");
//     console.log(bookmarkStart)
//     assert(8, bookmarkStart.children.length)
// });
//
// test('支持读取ffData里的checkbox的状态', async function() {
//     const ffData = await xmlreader.readString('<w:ffData><w:name w:val="name"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:maxLength w:val="20"/></w:textInput></w:ffData>')
//     console.log(ffData)
// })


test("complex fields", (function() {
    var uri = "http://example.com";

    var ffDataXml1 = new XmlElement("w:ffData", {}, [
        new XmlElement("w:name", {"w:val": "name"}, []),
        new XmlElement("w:enabled", {}, []),
        new XmlElement("w:calcOnExit", {"w:val": "0"}, []),
        new XmlElement("w:textInput", {}, [
            new XmlElement("w:maxLength", {"w:val": "20"}, [])
        ]),
    ])
    var ffDataXml2 = new XmlElement("w:ffData", {}, [
        new XmlElement("w:name", {"w:val": "CheckBox7"}, []),
        new XmlElement("w:enabled", {}, []),
        new XmlElement("w:calcOnExit", {"w:val": "0"}, []),
        new XmlElement("w:checkBox", {}, [
            new XmlElement("w:checked", {"w:val": "0"}, [])
        ]),
    ])
    var ffDataXml3 = new XmlElement("w:ffData", {}, [
        new XmlElement("w:name", {"w:val": "CheckBox7"}, []),
        new XmlElement("w:enabled", {}, []),
        new XmlElement("w:calcOnExit", {"w:val": "0"}, []),
        new XmlElement("w:checkBox", {}, [
            new XmlElement("w:checked", {}, [])
        ]),
    ])
    var ffDataXml4 = new XmlElement("w:ffData", {}, [
        new XmlElement("w:name", {"w:val": "ddList"}, []),
        new XmlElement("w:enabled", {}, []),
        new XmlElement("w:calcOnExit", {"w:val": "0"}, []),
        new XmlElement("w:ddList", {}, [
            new XmlElement("w:result", {"w:val": "1"}, []),
            new XmlElement("w:listEntry", {"w:val": "01"}, []),
            new XmlElement("w:listEntry", {"w:val": "02"}, []),
            new XmlElement("w:listEntry", {"w:val": "03"}, []),
            new XmlElement("w:listEntry", {"w:val": "04"}, []),
            new XmlElement("w:listEntry", {"w:val": "05"}, []),
        ]),
    ])
    var beginXml = new XmlElement("w:r", {}, [
        new XmlElement("w:fldChar", {"w:fldCharType": "begin"}, [ffDataXml1])
    ]);
    var endXml = new XmlElement("w:r", {}, [
        new XmlElement("w:fldChar", {"w:fldCharType": "end"})
    ]);
    var separateXml = new XmlElement("w:r", {}, [
        new XmlElement("w:fldChar", {"w:fldCharType": "separate"})
    ]);
    var inputValueXml = new XmlElement("w:r", {}, [
        runOfText("梁栓1"),
        runOfText("梁栓2"),
        runOfText("梁栓3"),
    ]);
    var hyperlinkInstrText = new XmlElement("w:instrText", {}, [
        xml.text('FORMTEXT')
    ]);
    var hyperlinkRunXml = runOfText("this is a hyperlink");

    var isEmptyHyperlinkedRun = isHyperlinkedRun({children: []});

    function isHyperlinkedRun(hyperlinkProperties) {
        return isRun({
            children: contains(
                isHyperlink(hyperlinkProperties)
            )
        });
    }


    return {

        "解析复杂的field": function() {

            const paragraphXml = new XmlElement("w:p", {}, [
                beginXml,
                hyperlinkInstrText,
                separateXml,
                inputValueXml,
                endXml
            ]);
            var paragraph = readXmlElementValue(paragraphXml);
            console.log(JSON.stringify(paragraph));

        },
        "解析多级的field": function() {

            var beginXml1 = new XmlElement("w:r", {}, [
                new XmlElement("w:fldChar", {"w:fldCharType": "begin"}, [ffDataXml1])
            ]);

            var beginXml2 = new XmlElement("w:r", {}, [
                new XmlElement("w:fldChar", {"w:fldCharType": "begin"}, [ffDataXml2])
            ]);
            const tcXml1 = new XmlElement("w:tc", {}, [
                new XmlElement("w:p", {}, [
                    beginXml1,
                    hyperlinkInstrText,
                    separateXml,
                    endXml
                ]),
            ]);
            const tcXml2 = new XmlElement("w:tc", {}, [
                new XmlElement("w:p", {}, [
                    beginXml2,
                    hyperlinkInstrText,
                    separateXml,
                    inputValueXml,
                    endXml
                ]),
            ])
            const trXml = new XmlElement("w:tr", {},[
                tcXml1,
                tcXml2
            ]);
            var tr = readXmlElementValue(trXml);
            console.log(JSON.stringify(tr));
        }
    };
})());


function createEmbeddedBlip(relationshipId) {
    return new XmlElement("a:blip", {"r:embed": relationshipId});
}

function createLinkedBlip(relationshipId) {
    return new XmlElement("a:blip", {"r:link": relationshipId});
}

function runOfText(text) {
    var textXml = new XmlElement("w:t", {}, [xml.text(text)]);
    return new XmlElement("w:r", {}, [textXml]);
}

function hyperlinkRelationship(relationshipId, target) {
    return {
        relationshipId: relationshipId,
        target: target,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    };
}

function imageRelationship(relationshipId, target) {
    return {
        relationshipId: relationshipId,
        target: target,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    };
}

function NumberingMap(options) {
    var findLevel = options.findLevel;
    var findLevelByParagraphStyleId = options.findLevelByParagraphStyleId || {};

    return {
        findLevel: function(numId, level) {
            return findLevel[numId][level];
        },
        findLevelByParagraphStyleId: function(styleId) {
            return findLevelByParagraphStyleId[styleId];
        }
    };
}
