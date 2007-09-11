//
// openOutlook.js - revision 1 (2007-09-11 / adam.nielsen@uq.edu.au)
//
// Read in the XML file given on the command line and use it to open a new
// preformatted e-mail in Outlook, ready to be sent.  Usage:
//
//   openOutlook.js data.xml
//
// Where data.xml looks like:
//
//   <mail>
//     <to>Adam Nielsen</to>
//     <to>John Smith</to>
//     <cc>Someone else</cc>
//     <subject>Test message</subject>
//     <body href="msg.html"/>
//   </mail>
//
// The message body can also be specified inline:
//
//     <body>Plain text message</body>
//
// And the XML file can also be specified in the strFilename variable below.
//
// Note that angle brackets still need to be escaped as &lt; and &gt; in the
// message body.
//
// This script is a text file, so make sure it uses DOS (CRLF) line endings or
// it will not run.
//
// This has been tested with Outlook 2003.
//

var strFilename = "$DATAFILE;";

// Script parameters
if (WScript.Arguments.Count() != 1) { // no parameter
	if (strFilename[0] == '$') { // no inline replacement
		// No data file supplied
		MessageBox("Usage: openOutlook data.xml", "Missing Parameter", MB_OK | MB_ICONEXCLAMATION);
		WScript.Quit();
	} // else continue on with inline replacement
} else {
	var strFilename = WScript.Arguments.Item(0);
}

// Globals
var olMailItem = 0; // TODO: check this
var olFormatHTML = 2;

// --- BEGIN MESSAGEBOX DEFINITION ---
var MB_OK = 0;
var MB_OKCANCEL = 1;
var MB_ABORTRETRYIGNORE = 2;
var MB_YESNOCANCEL = 3;
var MB_YESNO = 4;
var MB_RETRYCANCEL = 5;
var MB_ICONERROR = 16;
var MB_ICONQUESTION = 32;
var MB_ICONEXCLAMATION = 48;
var MB_ICONINFORMATION = 64;
var MB_FOCUS_BTN0 = 0;
var MB_FOCUS_BTN1 = 256;
var MB_FOCUS_BTN2 = 512;
var MB_RET_OK = 1;
var MB_RET_CANCEL = 2;
var MB_RET_ABORT = 3;
var MB_RET_RETRY = 4;
var MB_RET_IGNORE = 5;
var MB_RET_YES = 6;
var MB_RET_NO = 7;
function MessageBox(strMessage, strTitle, iButtons)
{
	var sh = WScript.CreateObject("WScript.Shell");
	sh.Popup(strMessage, 0, strTitle, iButtons);
}
// --- END MESSAGEBOX DEFINITION ---

//var ol = Script.CreateObject("Outlook.Application");

var doc = new ActiveXObject("msxml2.DOMDocument.3.0");

doc.async = false;
doc.resolveExternals = false;
doc.validateOnParse = false;

if (!doc.load(strFilename)) {
	MessageBox("Unable to load " + strFilename +
		"\n\nReason: " + doc.parseError.reason, "Error", MB_OK | MB_ICONERROR);
	WScript.Quit();
}

function joinNodes(nodeList, delim)
{
	if (nodeList.length == 0) return "";
	var str = nodeList[0].text;
	for (var i = 1; i < nodeList.length; i++) {
		str += delim + nodeList[i].text;
	}
	return str;
}

var ol = WScript.CreateObject("Outlook.Application");
var msg = ol.CreateItem(olMailItem);

var nodeTo = doc.selectNodes("//mail/to");
if (nodeTo) msg.To = joinNodes(nodeTo, "; "); // Semicolon-delimited list of recipients

var nodeCC = doc.selectNodes("//mail/cc");
if (nodeCC) msg.CC = joinNodes(nodeCC, "; ");

var nodeSubject = doc.selectSingleNode("//mail/subject");
if (nodeSubject) msg.Subject = nodeSubject.text;

var nodeBody = doc.selectSingleNode("//mail/body");
if (nodeBody) {
	if (nodeBody.text != "") msg.Body = nodeBody.text;
	else {
		var strBodyRef = nodeBody.getAttribute("href");
		if (strBodyRef != "") {
			// The body is stored in an external file, so load that

			var fso = new ActiveXObject("Scripting.FileSystemObject");
			var FOR_READING = 1, FOR_WRITING = 2, FOR_APPENDING = 8;

			var f = fso.OpenTextFile(strBodyRef, FOR_READING);
			var strHTML = f.ReadAll();
			f.close();

			msg.BodyFormat = olFormatHTML;
			msg.HTMLBody = strHTML;//"<code>" + strBodyRef + "</code>";
		}
	}
}

msg.Display();
