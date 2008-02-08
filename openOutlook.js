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
//     <bcc>Bob Jones</bcc>
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

// Globals
var olMailItem = 0; // TODO: check this
var olFormatHTML = 2;
var olByValue = 1;
var olSave = 0;
//var olEmbeddedItem = 5;
var CdoPR_ATTACH_MIME_TAG = 0x370E001E;

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
	if (strFilename[0] == '$') { // no inline replacement
		// No data file supplied
		MessageBox("Usage: openOutlook data.xml", "Missing Parameter", MB_OK | MB_ICONEXCLAMATION);
		WScript.Quit();
	} // else continue on with inline replacement
} else {
	var strFilename = WScript.Arguments.Item(0);
}

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

var nodeBCC = doc.selectNodes("//mail/bcc");
if (nodeBCC) msg.BCC = joinNodes(nodeBCC, "; ");

var nodeSubject = doc.selectSingleNode("//mail/subject");
if (nodeSubject) msg.Subject = nodeSubject.text;

var nodeAttachments = doc.selectNodes("//mail/attachment");
if (nodeAttachments.length > 0) {
	//var str = nodeAttachments[0].text;
	var strCIDs = new Array();
	for (var i = 0; i < nodeAttachments.length; i++) {
		var strDesc = nodeAttachments[i].getAttribute("description");
		if (strDesc == null) {
			MessageBox("<attachment/> tag has no description attribute!", "XML Error", MB_OK | MB_ICONERROR);
			WScript.Quit();
		}

		var att = msg.Attachments.Add(nodeAttachments[i].text, olByValue, 1, strDesc);
		att = null;
		strCIDs[i] = nodeAttachments[i].getAttribute("cid");
		//att.Position = -1;
		//alert(att.position);
		//att.Fields.add(0x7FFF000B, "true");
		//att.Fields.add(0x370B0003, -1);
	}
	// We have to release the attachment objects fully, otherwise the changes won't be saved
	CollectGarbage(); // release all the objects we've assigned null to

	//msg.Save();
	//var strMsgID = msg.EntryID;
	//alert(strMsgID);
	msg.Close(olSave);
	var strMsgID = msg.EntryID;

	// We have to release the objects fully, otherwise the changes won't be saved
	msg = null;
	ol = null;
	CollectGarbage(); // release all the objects we've assigned null to

	// Find the message in "MAPI Session" mode
	var oSession = WScript.CreateObject("MAPI.Session");
	try {
		oSession.Logon("", "", false, false);
	} catch (e) {
		if (e.number == -2147221231) {
			MessageBox("Unable to generate e-mail message.  Are you sure you have Outlook open?", "Outlook error", MB_OK | MB_ICONERROR);
		} else {
			throw e;
		}
		WScript.Quit();
	}

	var oMsg = oSession.GetMessage(strMsgID);

	// Set the attachment CIDs
	for (var i = 0; i < nodeAttachments.length; i++) {
		var oAttachFields = oMsg.Attachments.Item(i+1).Fields;
		//oAttachFields.Add(CdoPR_ATTACH_MIME_TAG, "image/jpeg");
		oAttachFields.Add(0x3712001E, strCIDs[i]);
	}

	// Hide the attachments (necessary?  Might be automatic for embedded images.)
	//oMsg.Fields.Add("{0820060000000000C000000000000046}0x8514", 11, true)

	// Save changes
	oMsg.Update();

	// Reopen the message in MAPI mode
	ol = WScript.CreateObject("Outlook.Application");
	// Get the Outlook MailItem again
	msg = ol.GetNamespace("MAPI").GetItemFromID(strMsgID);
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

			var f = fso.OpenTextFile(strBodyRef, FOR_READING);
			var strHTML = f.ReadAll();
			f.close();

			msg.BodyFormat = olFormatHTML;
			msg.HTMLBody = strHTML;//"<code>" + strBodyRef + "</code>";
		}
	}
}

msg.Close(olSave);

msg.Display();
