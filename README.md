# Google-Scripts-form-to-pdf

Let's imagine you have a process in your company that consists of a form that your customers should fill, print, sign and return to you. It could be a hassle to have just a word or pdf document to send to people as they might fill it in different ways, and most important, it's difficult to track who has filled it.

Ok, so the solution I came up with was to create a Google Form that would process the results, put them in a template and send to the user's e-mail in pdf format, ready to print.

You can test it here: <a href="https://docs.google.com/forms/d/1v73wh1E8B4eMUfQBIPSrQHNOWPN5nBnbLAmPO7MKxcY/viewform" target="_blank">TEST FORM</a>

_Implementation:

1. Create your Google Docs template (see the example in this repo <a href="https://docs.google.com/document/d/1F08ZphFd-KWl9q2C0FUbpBV7sLShv5aeOgkOCDPVpC0/edit?usp=sharing target="_blank"">here</a>) and fill it with the "wild cards" that will be replaced with the real info. Be careful with the keywords as the replacement function works for all the strings that matches what you are looking for!
2. Create your Google Form based on the answers you have in your template.
3. Open the Google Spreadsheet that was created along with your form and open Tools > Script editor...
4. Paste the script from this repo and change to fit your needs
5. Go to Resources > Current project's triggers and add a new trigger "From spreadsheet" "onFormSubmit"
6. Don't forget to run it once from the GoogleScripts platform for authorizing the execution (just need to be done once)

And that's it! Every time someone submits an answer he or she will get an e-mail in less than one minute with a pdf attachment, ready to print.
