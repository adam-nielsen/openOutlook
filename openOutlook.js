//
// openOutlook.js - revision 3 (2011-04-04 / adam.nielsen@uq.edu.au)
//
// Read in the XML file given on the command line and use it to open a new
// preformatted e-mail in Outlook, ready to be sent.  Supports full HTML
// content in the message body, with embedded images and other attachments.
//
// This file is placed in the public domain.  Feel free to use it for anything
// you like.  There is no warranty - if it breaks you get to keep both pieces.
//
// 2007-09-11 / adam.nielsen@uq.edu.au: Initial version
// 2009-07-07 / adam.nielsen@uq.edu.au: Set message body to UTF-8
// 2011-04-04 / adam.nielsen@uq.edu.au: Remove CDO code to hide image
//   attachments, seems no longer req'd and now works with Outlook 2010
//
// Usage:
//
//   openOutlook.js data.xml
//
// Where data.xml looks like:
//
//   <mail>
//     <to>Adam Nielsen</to>
//     <to>John Smith</to>
//     <cc>Someone else</cc>
//     <bcc>Bob Jones</bcc>
//     <subject>Test message</subject>
//     <body href="msg.html"/>
//     <attachment cid="cid:1" href="example.jpg">Embedded image</attachment>
//   </mail>
//
// All fields are optional.  The "cid" can be used with attachments to link to
// embedded images from within the message body, like <img src="cid:1" />
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
// This has been tested with Outlook 2007 and Outlook 2010.
//

// If you don't want to specify the XML filename on the command line, you can
// hard code it here.  The reason for it saying "datafile" at the moment is
// because in our environment JavaScript code running through XULRunner
// overwrites this placeholder with the path to the XML file before launching
// the script.
var strFilename = "$DATAFILE;";

// Globals
var olMailItem = 0;
var olFormatHTML = 2;
var olByValue = 1;
var olSave = 0;

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

// Script parameters
if (WScript.Arguments.Count() != 1) { // no parameter
	if (strFilename.substr(0, 1) == '$') { // no inline replacement
		// No data file supplied
		MessageBox("Usage: openOutlook data.xml", "Missing Parameter", MB_OK | MB_ICONEXCLAMATION);
		WScript.Quit();
	} // else continue on with inline replacement
} else {
	strFilename = WScript.Arguments.Item(0);
}

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
msg.InternetCodePage = 65001; // UTF-8

var nodeTo = doc.selectNodes("//mail/to");
if (nodeTo) msg.To = joinNodes(nodeTo, "; "); // Semicolon-delimited list of recipients

var nodeCC = doc.selectNodes("//mail/cc");
if (nodeCC) msg.CC = joinNodes(nodeCC, "; ");

var nodeBCC = doc.selectNodes("//mail/bcc");
if (nodeBCC) msg.BCC = joinNodes(nodeBCC, "; ");

var nodeSubject = doc.selectSingleNode("//mail/subject");
if (nodeSubject) msg.Subject = nodeSubject.text;

var nodeAttachments = doc.selectNodes("//mail/attachment");
if (nodeAttachments.length > 0) {
	var strCIDs = new Array();
	for (var i = 0; i < nodeAttachments.length; i++) {
		var strHRef = nodeAttachments[i].getAttribute("href");
		if (strHRef == null) {
			MessageBox("<attachment/> tag has no href attribute!", "XML Error", MB_OK | MB_ICONERROR);
			WScript.Quit();
		}

		var att = msg.Attachments.Add(strHRef, olByValue, 1, nodeAttachments[i].text);
		att = null; // necessary! (for garbage collection)
		strCIDs[i] = nodeAttachments[i].getAttribute("cid");
	}
	// We have to release the attachment objects fully, otherwise the changes won't be saved
	CollectGarbage(); // release all the objects we've assigned null to
}

var nodeBody = doc.selectSingleNode("//mail/body");
if (nodeBody) {
	if (nodeBody.text != "") msg.Body = nodeBody.text;
	else {
		var strBodyRef = nodeBody.getAttribute("href");
		if (strBodyRef != "") {
			// The body is stored in an external file, so load that

			var fso = new ActiveXObject("Scripting.FileSystemObject");
			var FOR_READING = 1, FOR_WRITING = 2, FOR_APPENDING = 8;

			var f = fso.OpenTextFile(strBodyRef, FOR_READING); // US-ASCII
			var strHTML = f.ReadAll();
			f.close();

			msg.BodyFormat = olFormatHTML;
			msg.HTMLBody = strHTML; // "<code>" + strBodyRef + "</code>";
		}
	}
}

msg.Close(olSave);

msg.Display();
