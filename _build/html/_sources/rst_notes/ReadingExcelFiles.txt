Reading Excel Files
===================

All the examples shown below can be found in the
xlrd
directory of the course material.

Opening Workbooks
-----------------

Workbooks can be loaded either from a file, an
mmap.mmap
object or from a string:

Navigating a Workbook
---------------------

Here is a simple example of workbook navigation:

The next few sections will cover the navigation of workbooks in more detail.




Introspecting a Book
~~~~~~~~~~~~~~~~~~~~

The
xlrd.Book
object returned by
open_workbook
contains all information to do with the workbook and can be used to retrieve individual sheets within the workbook.

The
nsheets
attribute is an integer containing the number of sheets in the workbook. This attribute, in combination with the
sheet_by_index

method, is the most common way of retrieving individual sheets.

The
sheet_names
method returns a list of unicodes containing the names of all sheets in the workbook. Individual sheets can be retrieved using these names by way of the
sheet_by_name
function.

The results of the
sheets
method can be iterated over to retrieve each of the sheets in the workbook.

The following example demonstrates these methods and attributes:

xlrd.Book
objects have other attributes relating to the content of the workbook that are only rarely useful:

* codepage


* countries


* user_name


If you think you may need to use these attributes, please see the
xlrd
documentation.




Introspecting a Sheet
~~~~~~~~~~~~~~~~~~~~~~

The
xlrd.sheet.Sheet
objects returned by any of the methods described above contain all the information to do with a worksheet and its contents.

The
name
attribute is a unicode representing the name of the worksheet.

The
nrows
and
ncols
attributes contain the number of rows and the number of columns, respectively, in the worksheet.

The following example shows how these can be used to iterate over and display the contents of one worksheet:

xlrd.sheet.Sheet
objects have other attributes relating to the content of the worksheet that are only rarely useful:

* col_label_ranges


* row_label_ranges


* visibility


If you think you may need to use these attributes, please see the
xlrd
documentation.




Getting a particular Cell
~~~~~~~~~~~~~~~~~~~~~~~~~

As already seen in previous examples, the
cell
method of a
Sheet
object can be used to return the contents of a particular cell.

The
cell
method returns an
xlrd.sheet.Cell
object. These objects have very few attributes, of which
value
contains the actual value of the cell and
ctype
contains the type of the cell.

In addition,
Sheet
objects have two methods for returning these two types of data. The
cell_value
method returns the value for a particular cell, while the
cell_type
method returns the type of a particular cell. These methods can be quicker to execute than retrieving the
Cell
object.

Cell types are covered in more detail later. The following example shows the methods, attributes and classes in action:





Iterating over the contents of a Sheet
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

We've already seen how to iterate over the contents of a worksheet and retrieve the resulting individual cells. However, there are methods to retrieve groups of cells more easily. There are a symmetrical set of methods that retrieve groups of cell information either by row or by column.

The
row
and
col
methods return all the
Cell
objects for a whole row or column, respectively.

The
row_slice
and
col_slice
methods return a list of
Cell
objects in a row or column, respectively, bounded by and start index and an optional end index.

The
row_types
and
col_types
methods return a list of integers representing the cell types in a row or column, respectively, bounded by and start index and an optional end index.

The
row_values
and
col_values
methods return a list of objects representing the cell values in a row or column, respectively, bounded by a start index and an optional end index.

The following examples demonstrates all of the sheet iteration methods:




Utility Functions
~~~~~~~~~~~~~~~~~

When navigating around a workbook, it's often useful to be able to convert between row and column indexes and the Excel cell references that users may be used to seeing. The following functions are provided to help with this:

The
cellname
function turns a row and column index into a relative Excel cell reference.

The
cellnameabs
function turns a row and column index into
an absolute Excel cell reference.

The
colname
function turns a column index into an Excel column name.

These three functions are demonstrated in the following example:

Unicode
-------

All text attributes and values produced by
xlrd
will be either unicode objects or, in rare cases, ascii strings.

Each piece of text in an Excel file written by Microsoft Excel is encoded into one of the following:

* Latin1, if it fits


* UTF_16_LE, if it doesn't fit into Latin1


* In older files, by an encoding specified by an MS codepage. These are mapped to Python encodings by
  xlrd
  and still result in unicode objects.


In rare cases, other software has been known to write no codepage or the wrong codepage into Excel files. In this case, the correct encoding may need to be specified to
open_workbook
:



Types of Cell
-------------

We have already seen the cell type expressed as an integer. This integer corresponds to a set of constants in xlrd that identify the type of the cell. The full set of possible cell types is listed in the following sections.




Text
~~~~

These are represented by the
xlrd.XL_CELL_TEXT
constant.

Cells of this type will have values that are
unicode
objects.

Number
~~~~~~

These are represented by the
xlrd.XL_CELL_NUMBER
constant.

Cells of this type will have values that are
float
objects.

Date
~~~~

These are represented by the
xlrd.XL_CELL_DATE
constant.

NB:
Dates don't really exist in Excel files, they are merely Numbers with a particular number formatting.

xlrd
will return
xlrd.XL_CELL_DATE
as the cell type if the number format string looks like a date.

The
xldate_as_tuple
method is provided for turning the
float
in a Date cell into a tuple suitable for instantiating various date/time objects. This example shows how to use it:

Caveats:

* Excel files have two possible date modes, one for files originally created on Windows and one for files originally created on an Apple machine. This is expressed as the
  datemode
  attribute of
  xlrd.Book
  objects and
  must
  be passed to xldate_as_tuple.


* The Excel file format has various problems with dates before 3 Jan 1904 that can cause date ambiguities that can result in
  xldate_as_tuple
  raising an XLDateError.


* The Excel
  formula function
  DATE()
  can return unexpected dates in certain circumstances.


Boolean
~~~~~~~

These are represented by the
xlrd.XL_CELL_BOOLEAN
constant.

Cells of this type will have values that are
bool
objects.

Error
~~~~~

These are represented by the
xlrd.XL_CELL_ERROR
constant.

Cells of this type will have values that are integers representing specific error codes.

The
error_text_from_code
dictionary can be used to turn error codes into error messages:

For a simpler way of sensibly displaying all cell types, see
xlutils.display
.




Empty / Blank
~~~~~~~~~~~~~

Excel files only store cells that either have information in them or have formatting applied to them. However,
xlrd
presents sheets as rectangular grids of cells.

Cells where no information is present in the Excel file are represented by the
xlrd.XL_CELL_EMPTY
constant. In addition, there is only one “empty cell”, whose value is an empty string, used by
xlrd
, so empty cells may be checked using a Python identity check.

Cells where only formatting information is present in the Excel file are represented by the
xlrd.XL_CELL_BLANK
constant and their value will always be an empty string.


The following example brings all of the above cell types together and shows examples of their use:


Names
-----

These are an infrequently used but powerful way of abstracting commonly used information found within Excel files.

They have many uses, and
xlrd
can extract information from many of them. A notable exception are names that refer to sheet and VBA macros, which are extracted but should be ignored.

Names are created in Excel by navigating to
Insert > Name > Define
. If you plan to use
xlrd
to extract information from Names, familiarity with the definition and use of names in your chosen spreadsheet application is a good idea.

Types
~~~~~

A Name can refer to:

* A constant

  * CurrentInterestRate = 0.015


  * NameOfPHB = “Attila T. Hun”



* An absolute (i.e. not relative) cell reference

  * CurrentInterestRate = Sheet1!$B$4



* Absolute reference to a 1D, 2D, or 3D block of cells

  * MonthlySalesByRegion = Sheet2:Sheet5!$A$2:$M$100



* A list of absolute references

  * Print
    _Titles =
    [row_header_ref, col_header_ref])



Constants can be extracted.

The coordinates of an absolute
reference can be extracted so that you can then extract the corresponding data from the relevant sheet(s).

A relative reference is useful only if you have external knowledge of what cells can be used as the origin. Many formulas found in Excel files include function calls and multiple references and are not useful, and can be too hard to evaluate.

A full calculation engine is not included in xlrd.

Scope
~~~~~

The scope of a Name can be global, or it may be specific to a particular sheet. A Name's identifier may be re-used in different scopes. When there are multiple Names with the same identifier, the most appropriate one is used based on scope. A good example of this is the built-in name Print_Area; each worksheet may have one of these.

Examples:

name=rate, scope=Sheet1, formula=0.015

name=rate, scope=Sheet2, formula=0.023

name=rate, scope=
*global*
, formula=0.040

A cell formula
(1+rate)^20
is equivalent to
1.015^20
if it appears in
Sheet1
but equivalent to
1.023^20
if it appears in
Sheet2
, and
1.040^20
if it appears in any other sheet.

Usage
~~~~~

Common reasons for using names include:

* Assigning textual names to values that may occur in many places within a workbook

  * eg:
    RATE = 0.015



* Assigning textual names to complex formulae that may be easily mis-copied

  * eg:
    SALES_RESULTS = $A$10:$M$999



Here's an example real-world use case: reporting to head office. A company's head office makes up a template workbook. Each department gets a copy to fill in. The various ranges of data to be provided all have defined names. When the files come back, a script is used to
validate that the department hasn't trashed the workbook and the names are used to extract the data for further processing. Using names decouples any artistic repositioning of the ranges, by either head office template-designing user or by departmental users who are filling in the template, from the script which only has to know what the names of the ranges are.

In the examples directory of the
xlrd
distribution you will find
namesdemo.xls
which has examples of most of the non-macro varieties of defined names. There is also
xlrdnamesAPIdemo.py
which shows how to use the name lookup dictionaries, and how to extract constants and references and the data that references point to.

Formatting
----------

We've already seen that
open_workbook
has a parameter to load formatting information from Excel files. When this is done, all the formatting information is available, but the details of how it is presented are beyond the scope of this tutorial.

If you wish to copy existing formatted data to a new Excel file, see
xlutils.copy
and
xlutils.filter
.

If you do wish to inspect formatting information, you'll need to consult the following attributes of the following classes:

xlrd.Book
~~~~~~~~~

xlrd.sheet.Sheet
~~~~~~~~~~~~~~~~

xlrd.sheet.Cell
~~~~~~~~~~~~~~~

xf_index

Other Classes
~~~~~~~~~~~~~

In addition, the following classes are solely used to represent formatting information:

Working with large Excel files
------------------------------

If you're working with particularly large Excel files then there are two features of xlrd that you should be aware of:

* The on_demand parameter can be passed as True to open_workbook resulting in worksheets only being loaded into memory when they are requested.


* xlrd.Book objects have an unload_sheet method that will unload worksheet, specified by either sheet index or sheet name, from memory.


The following example shows how a large workbook could be iterated over when only sheets matching a certain pattern need to be inspected, and where only one of those sheets ends up in memory at any one time:





Introspecting Excel files with runxlrd.py
-----------------------------------------

The xlrd source distribution includes a
runxlrd.py
script that is extremely useful for introspecting Excel files without writing a single line of Python.

You are encouraged to run a variety of the commands it provides over the Excel files provided in the course materials.

The following gives an overview of what's available from
runxlrd
, and can be obtained using
python runxlrd.py –-help
:

runxlrd.py [options] command [input-file-patterns]


Commands:


2rows
Print the contents of first and last row in each sheet

3rows
Print the contents of first, second and last row in each sheet

bench
Same as "show", but doesn't print -- for profiling

biff_count[1]
Print a count of each type of BIFF record in the file

biff_dump[1]
Print a dump (char and hex) of the BIFF records in the file

fonts
hdr + print a dump of all font objects

hdr
Mini-overview of file (no per-sheet information)

hotshot
Do a hotshot profile run e.g. ... -f1 hotshot bench bigfile*.xls

labels
Dump of sheet.col_label_ranges and ...row... for each sheet

name_dump
Dump of each object in book.name_obj_list

names
Print brief information for each NAME record

ov
Overview of file

profile
Like "hotshot", but uses cProfile

show
Print the contents of all rows in each sheet

version[0]
Print versions of xlrd and Python and exit

xfc
Print "XF counts" and cell-type counts -- see code for details


[0] means no file arg

[1] means only one file arg i.e. no glob.glob pattern



Options:

-h, --help
show this help message and exit

-l LOGFILENAME, --logfilename=LOGFILENAME

contains error messages

-v VERBOSITY, --verbosity=VERBOSITY

level of information and diagnostics provided

-p PICKLEABLE, --pickleable=PICKLEABLE

1: ensure Book object is pickleable (default); 0:

don't bother

-m MMAP, --mmap=MMAP
1: use mmap; 0: don't use mmap; -1: accept heuristic

-e ENCODING, --encoding=ENCODING

encoding override

-f FORMATTING, --formatting=FORMATTING

0 (default): no fmt info 1: fmt info (all cells)

-g GC, --gc=GC
0: auto gc enabled; 1: auto gc disabled, manual

collect after each file; 2: no gc

-s ONESHEET, --onesheet=ONESHEET

restrict output to this sheet (name or index)

-u, --unnumbered
omit line numbers or offsets in biff_dump

