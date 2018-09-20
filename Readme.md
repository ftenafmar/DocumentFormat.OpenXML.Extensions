## Overview
 
Ported to NetStandard 2.0 the original library from James Westgate with some minor fixes (https://simpleooxml.codeplex.com).

Simple OOXML makes the creation of Open Office XML documents easier for developers. Modify or create any .docx or .xlsx document without Microsoft Word or Microsoft Excel.  Uses the Open Office SDK v 2.0.

_(Please download and install the Open XML Format SDK v 2.0 to use this library at [http://www.microsoft.com/downloads/details.aspx?FamilyId=C6E744E5-36E9-45F5-8D8C-331DF206E0D0](http://www.microsoft.com/downloads/details.aspx?FamilyId=C6E744E5-36E9-45F5-8D8C-331DF206E0D0))_

The goal of this project is the simple, effective creation of documents and spreadsheets using minimum resources, including a server environment. The library provides commonly used functionality whilst hides away the details of creating open xml documents and without a large performance overhead. Documents created with this library and the Open Office SDK can be viewed using Microsoft Excel/Microsoft Word or OpenOffice as well as any third party that supports the format.

## Getting Started

Simple OOXML adds the _DocumentFormat.OpenXml.Extensions_ namespace to version 2.0 of the Open Office SDK. It allows developers to create spreadsheets and documents either from scratch or using predefined templates. All functionality is represented by static functions for high performance tasks, or higher level wrapper functions can provide simpler code expressions with some minor performance loss.  

The following classes are provided:
* SpreadsheetReader - manipulation of templates, retrieval of document parts, row and column reference functionality
* SpreadsheetWriter - writing of document parts and creation of document level attributes. Add or remove spreadsheets.
* SpreadsheetStyle - encapsulates font, border and fill handing in a spreadsheet.
* WoksheetReader - retrieves cell and style information from a worksheet
* WorksheetWriter - allows the pasting or insertion of data and style - using simple value types or DataTables - at a cell or range reference.
* DocumentReader - retrieval of document templates.
* DocumentWriter - pastes and saves text and text lists using predefined bookmarks.

Download the source files to view the source code, examples as well as a unit testing library which is a useful reference to all the features of the library. Users of the unsupported ExcelPackage library could consider using this library instead.

Simple OOXML is licensed under the permissive MIT licence. 