# Project Description
Yet another wrapper around the open xml office package.
You can easily generate xlsx documents based on a template xlsx document, i.e. reuse parts from that template document, by marking them as named ranges (i.e."names" in Office terminology).
You can even reuse entire sheets as templates, thus the sky is the limit.

Basically you design your document normally in Excel, you save it and you can start create new xls files based on this file programatically. You will be able to reuse different parts from the original document in many places in the one that you are creating (usually these parts you will place in a separate sheet and you will delete them at the end). 
The granularity of reusable parts spans from one cell to range of cells and even entire sheets.

The Excel documents, which you are using as templates, can be altered later and in case you are not contradicting with your prevous logic (i.e. delete the used named ranges, etc.), your c# code (or whatever language you are using) doesn't have to be changed at all. This is great, when you wish to have multilingual support or different themes.

For a sample code, how to use it, see [Documentation], or check the folder "sqliteexporter" of the source code for a complete(!) project example.


Requirement: .Net 3.5 or later.
Microsoft Office does not need to be installed!

Just one tipp:
Even earlier versions of Office applications (2003 or 2K) can open the new xlsx files, if you install this ("Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint File Formats"):
http://www.microsoft.com/downloads/en/details.aspx?familyid=941b3470-3ae9-4aee-8f43-c6bb74cd1466&displaylang=en

[//]: #

   [Documentation]: <https://github.com/asinoai/officehelper/blob/master/documentation.md>
