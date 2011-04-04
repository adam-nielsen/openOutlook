## openOutlook.js ##
Written by Adam Nielsen <<adam.nielsen@uq.edu.au>>
https://github.com/adam-nielsen/openOutlook

This JScript code reads in the XML file given on the command line and uses it to open a new preformatted e-mail in Outlook, ready to be sent.  It supports full HTML content in the message body, with embedded images and other attachments.

This file is placed in the public domain.  Feel free to use it for anything you like.  There is no warranty - if it breaks you get to keep both pieces.

Usage:

    wscript openOutlook.js data.xml

Where data.xml looks like:

    <mail>
      <to>Adam Nielsen</to>
      <to>John Smith</to>
      <cc>Someone else</cc>
      <bcc>Bob Jones</bcc>
      <subject>Test message</subject>
      <body href="msg.html"/>
      <attachment cid="cid:1" href="example.jpg">Embedded image</attachment>
    </mail>

All fields are optional.  The "cid" can be used with attachments to link to embedded images from within the message body, like `<img src="cid:1" />`

The message body can also be specified inline:

    <body>Plain text message</body>

And the XML file can also be specified in the strFilename variable below.

Note that angle brackets still need to be escaped as &lt; and &gt; in the message body.

This script is a text file, so make sure it uses DOS (CRLF) line endings or it will not run.

This has been tested with Outlook 2007 and Outlook 2010.  Earlier revisions worked better with Outlook 2003.
