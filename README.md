# Presentation Text Replacer
Replaces text inside ppt, pptx and compatible formats.
Very useful to automate presentations and keep them always updated!

Just call using the paremeters explained in the arguments.py file.

You can create a presentation on Google Spreadsheets and set text tokens on this presentation, i.e: {total_sales}.

Then, using any automated script, you create a .ini file like this:

```
[replaces]
total_sales = 10.5M
```

Run this script over the presentation and all tokens will be replaced, updating the presentation without change its formatting.

Trust me: when you have dozens of contracts of different companies, media kits, proposals and other files to keep updated, automation is the way out to avoid losing deadlines and hairs.

Create a .ini file is simple and you can do it using a variety of tools. This script sends info from ini file to a pptx or ppt file without you even have to open it in PowerPoint (or similar).
