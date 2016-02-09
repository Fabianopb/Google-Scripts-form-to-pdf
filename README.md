# Google-Scripts-form-to-pdf

Let's imagine you have a process in your company that consists of a form that your customers should fill, print, sign and return to you. It could be a hassle to have just a word or pdf document to send to people as they might fill it in different ways, and most important, it's difficult to track who has filled it.

Ok, so the solution I came up with was to create a Google Form that would process the results, put them in a template and send to the user's e-mail in pdf format, ready to print.

You can test it here: <ADD LINK LATER>

_Implementation:

1. Create your Google Docs template and fill it with the "wild cards" that will be replaced with the real info. Be careful with the keywords as the replacement function works for all the strings that matches what you are looking for!
2. Create your Google Form based on the answers you have in your template.
3. Open the Google Spreadsheet that was created along with your form and open Tools > Script editor...
4. Paste the script from this repo and change to fit your needs

And that's it! Every time someone submits an answer he or she will get an e-mail in less than one minute with a pdf attachment, ready to print.
