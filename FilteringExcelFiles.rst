Filtering Excel Files
=====================

Any examples shown below can be found in the 'xlutils' directory of the course material.

Other utilities in xlutils
--------------------------

The 'xlutils' package contains several utilities in addition to those for filtering. The following are often useful:

xlutils.styles
~~~~~~~~~~~~~~

This module contains one class which, when instantiated with an 'xlrd.Workbook', will let you discover the style name and information from a given cell in that workbook as shown in the following example:

::

  from xlrd import open_workbook
  from xlutils.styles import Styles
  
  book = open_workbook('source.xls',formatting_info=True)
  styles = Styles(book)
  sheet = book.sheet_by_index(0)
  
  print styles[sheet.cell(1,1)].name
  print styles[sheet.cell(1,2)].name
  
  A1_style = styles[sheet.cell(0,0)]
  A1_font = book.font_list[A1_style.xf.font_index]
  print book.colour_map[A1_font.colour_index]
  
  styles.py

*NB:* For obvious reasons, 'open_workbook' must be called with 'formatting_info=True' in order to use 'xlutils.styles'.

Full documentation and examples can be found in the 'styles.txt' file in the docs folder of the 'xlutils' source distribution.

xlutils.display
~~~~~~~~~~~~~~~~

This module contains utility functions for easy and safe display of information returned by 'xlrd'.

'quoted_sheet_name' is called with the 'name' attribute of an 'xlrd.sheet.Sheet' instance and will return an encoded string containing a quoted version of the sheet's name.

'cell_display' is called with an 'xlrd.sheet.Cell' instance and returns an encoded string containing a sensible representation of the cells contents, even for Date and Error cells. If a date cell is to be displayed, 'cell_display' *must* be called with the 'datemode' attribute of the 'xlrd.Book' from which the cell came.

The following examples show both functions in action:

::

  from xlrd import open_workbook
  from xlutils.display import quoted_sheet_name
  from xlutils.display import cell_display
  
  wb = open_workbook('source.xls')
  
  print quoted_sheet_name(wb.sheet_names()[0])
  print repr(quoted_sheet_name(u'Price(\xa3)','utf-8'))
  print quoted_sheet_name(u'My Sheet')
  print quoted_sheet_name(u"John's Sheet")
  
  sheet = wb.sheet_by_index(0)
  print cell_display(sheet.cell(1,1))
  print cell_display(sheet.cell(1,3),wb.datemode)
  
  display.py

Full documentation and examples can be found in the 'display.txt' file in the docs folder of the 'xlutils' source distribution.

xlutils.copy
~~~~~~~~~~~~

This module contains one function that will take an 'xlrd.Book' and returns an 'xlwt.Workbook' populated with the data and formatting found in the 'xlrd.Book'.

This is extremely useful for updating an existing spreadsheet as the following example shows:

::

  from xlrd import open_workbook
  from xlwt import easyxf
  from xlutils.copy import copy
  
  rb = open_workbook('source.xls',formatting_info=True)
  rs = rb.sheet_by_index(0)
  wb = copy(rb)
  ws = wb.get_sheet(0)
  
  plain = easyxf('')
  for i,cell in enumerate(rs.col(2)):
      if not i:
          continue
      ws.write(i,2,cell.value,plain)
  
  for i,cell in enumerate(rs.col(4)):
      if not i:
          continue
      ws.write(i,4,cell.value-1000)
  
  wb.save('output.xls')
  
  copy.py

It is important to note that some things won't be copied:

* Formulae

* Names

* Anything ignored by 'xlrd'

In addition to the modules described above, there are also 'xlutils.margins' and 'xlutils.save', but these are only useful in certain situations. Refer to their documentation in the 'xlutils'source distribution.

Structure of xlutils.filter
---------------------------

This framework is designed to filter and split Excel files using a series of modular readers, filters and writers as shown in the diagram below:

.. image:: images/Object_1.png

The flow of information between the components is by method calls on the next component in the chain. The possible method calls are listed in the table below, where 'rdbook' is an 'xlrd.Book' instance; 'rdsheet' is an 'xlrd.sheet.Sheet' instance; 'rdrowx', 'rdcolx', 'wtrowx', and 'wtcolx' and integer indices specifying the cell to read from and write to; 'wtbook_name' is a string specifying the name of the Excel file to write to; and 'wtsheet_name' is a 'unicode' specifying the name of the sheet to write to:

+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| start()                           | This method is called before processing of a batch of input. It can be called at any time. One common use is to reset all the      |
|                                   | filters in a chain in the event of an error during the processing of an 'rdbook'.                                                  |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| workbook(rdbook,wtbook_name)      | This method is called every time processing of a new workbook starts                                                               |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| sheet(rdsheet,wtsheet_name)       | This method is called every time processing of a new sheet in the current workbook starts                                          |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| set_rdsheet(rdsheet)              | This method is called to indicate a change for the source of cells mid-way through writing a sheet.                                |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| row(rdrowx,wtrowx)                | The row method is called every time processing of a new row in the current sheet starts.                                           |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| cell(rdrowx,rdcolx,wtrowx,wtcolx) | This is called for every cell in the sheet being processed. This is the most common method in which filtering and queuing of onward|
|                                   | calls to the next component takes place.                                                                                           |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+
| finish                            | This method is called once processing of all workbooks has been completed.                                                         |
|                                   |                                                                                                                                    |
+-----------------------------------+------------------------------------------------------------------------------------------------------------------------------------+

Readers
~~~~~~~

A reader's job is to obtain one or more 'xlrd.Book' objects and iterate over those objects issuing appropriate calls to the next component in the chain. The order of calling is expected to be as follows:

* 'start'

  * 'workbook', once for each 'xlrd.Book' object obtained

    * 'sheet', once for each sheet found in the current book

    * 'set_rdsheet', whenever the sheet from which cells to be read needs to be changed. This method may not be called between calls to 'row' and 'cell', and between multiple calls to 'cell'. It may only be called once all 'cell' calls for a row have been made.

      * 'row', once for each row in the current sheet

        * 'cell', once for each cell in the row

* 'finish', once all 'xlrd.Book' objects have been processed

Also, for method calls made by a reader, the following should be true:

* 'wtbook_name' should be the filename of the file the 'xlrd.Book' object originated from.

* 'wtsheet_name' should be 'rdbook.name'

* 'wtrowx' should be equal to 'rdrowx'

* 'rdcolx' should be equal to 'wtcolx'

Because of these restrictions, an 'xlutils.filter.BaseReader' class is provided that will normally only need to have one of two methods overridden to get any required functionality:

* 'get_filepaths' – if implemented, this must return an iterable sequence of paths to excel files that can be opened with python's builtin file.

* 'get_workbooks' – if implemented, this must return an sequence of 2-tuples. Each tuple must contain an 'xlrd.Book' object followed by a string containing the filename of the file from which the 'xlrd.Book' object was loaded.

Filters
~~~~~~~

Implementing these components is where the bulk of the work will be done by users of the 'xlutils.filter' framework. A Filter's responsibilities are to accept method calls from the preceding component in the chain, do any processing necessary and then emit appropriate method calls to the next component in the chain.

There is very little constraint on what order Filters receive and emit method calls other than that the order of method calls emitted must remain consistent with the structure given above. This enables components to be freely interchanged more easily.

Because Filters may only need to implement few of the full set of method calls, an 'xlutils.filter.BaseFilter' is provided that does nothing but pass the method calls on to the next component in the chain. The implementation of this filter is useful to see when embarking on Filter implementation:

::

  class BaseFilter:
  
      def start(self):
          self.next.start()
  
      def workbook(self,rdbook,wtbook_name):
          self.next.workbook(rdbook,wtbook_name)
  
      def sheet(self,rdsheet,wtsheet_name):
          self.rdsheet = rdsheet
          self.next.sheet(rdsheet,wtsheet_name)
  
      def set_rdsheet(self,rdsheet):
          self.rdsheet = rdsheet
          self.next.set_rdsheet(rdsheet)
  
      def row(self,rdrowx,wtrowx):
          self.next.row(rdrowx,wtrowx)
  
      def cell(self,rdrowx,rdcolx,wtrowx,wtcolx):
          self.next.cell(rdrowx,rdcolx,wtrowx,wtcolx)
  
      def finish(self):
          self.next.finish()


Writers
~~~~~~~

These components do the grunt work of actually copying the appropriate information from the 'rdbook' and serialising it into an Excel file. This is a complicated process and not for the feint of hard to re-implement.

For this reason, an 'xlutils.filter.BaseWriter' component is provided that does all of the hard work and has one method that needs to be implemented. That method is 'get_stream' and it is called with the filename of the Excel file to be written. Implementations of this method are expected to return a new file-like object that has a 'write' and, by default, a 'close' method each time they are called.

Subclasses may also override the boolean 'close_after_write' attribute, which is 'True' by default, to indicate that the file-like objects returned from 'get_stream' should not have their 'close' method called once serialisation of the Excel file data is complete.

It is important to note that some things won't be copied from the 'rdbook' by 'BaseWriter':

* Formulae

* Names

* Anything ignored by 'xlrd'

Process
~~~~~~~

The process function is responsible for taking a series of components as its arguments. The first of these should be a Reader. The last of these should be a Writer. The rest should be the necessary Filters in the order of processing required.

The process method will wire these components together by way of their 'next' attributes and then kick the process off by calling the Reader and passing the first Filter in the chain as its argument.

A worked example
----------------

Suppose we want to filter an existing Excel file to omit rows that have an X in the first column.

The following example shows possible components to do this and shows how they would be instantiated and called to achieve this:

::

  import os
  from xlutils.filter import \ 
      BaseReader,BaseFilter,BaseWriter,process
  
  class Reader(BaseReader):
      def get_filepaths(self):
          return [os.path.abspath('source.xls')]
  
  class Writer(BaseWriter):
      def get_stream(self,filename):
          return file(filename,'wb')
  
  class Filter(BaseFilter):
      pending_row = None
      wtrowxi = 0
      def workbook(self,rdbook,wtbook_name):
          self.next.workbook(rdbook,'filtered-'+wtbook_name)
      def row(self,rdrowx,wtrowx):
          self.pending_row = (rdrowx,wtrowx)
      def cell(self,rdrowx,rdcolx,wtrowx,wtcolx):
          if rdcolx==0:
              value = self.rdsheet.cell(rdrowx,rdcolx).value
              if value.strip().lower()=='x':
                  self.ignore_row = True
                  self.wtrowxi -= 1
              else:
                  self.ignore_row = False
                  rdrowx, wtrowx = self.pending_row
                  self.next.row(rdrowx,wtrowx+self.wtrowxi)
          elif not self.ignore_row:
              self.next.cell(
                  rdrowx,rdcolx,wtrowx+self.wtrowxi,wtcolx-1
                  )        
  
  process(Reader(),Filter(),Writer())
  
  filter.py

In reality, we would not need to implement the Reader and Writer components, as there are already suitable components included.

Existing components
-------------------

The 'xlutils.filter' framework comes with a wide range of existing components, each of which is briefly described below. For full descriptions and worked examples of all these components, please see 'filter.txt' in the 'docs' folder of the 'xlutils' source distribution.

GlobReader
~~~~~~~~~~

If you're processing files that are on disk, then this is probably the reader for you. It returns all files matching the path specification it's instantiated with.

XLRDReader
~~~~~~~~~~

This reader can be used at the start of a chain when you already have an 'xlrd.Book' object and you'll looking to process it with 'xlutils.filter'.

TestReader
~~~~~~~~~~

This reader is specifically designed for testing filterimplementations with known sets of cells.

DirectoryWriter
~~~~~~~~~~~~~~~

If you want files you're processing to end up on disk, then this is probably the writer for you. It stores files in the directory it is instantiated with.

StreamWriter
~~~~~~~~~~~~

If you want to write exactly one workbook to a stream, such as a 'tempfile.TemporaryFile' or 'sys.stdout', then this is the writer for you.

XLWTWriter
~~~~~~~~~~

If you want to change cells after the filtering process is complete then this writer can be used to obtain the 'xlwt.Workbook' objects that BaseWriter generates.

ColumnTrimmer
~~~~~~~~~~~~~

This filter will strip columns containing no useful data from the end of sheets. The definition of “no useful data” can be controlled during instantiation of this filter.

ErrorFilter
~~~~~~~~~~~

This filter caches all method calls in a file on disk and will only pass them on the next component in the chain when its 'finish' method has been called '???' and no error messages have been logged to the python logging framework.

If Boolean or error Cells are encountered, an error message will be logged to the python logging framework will will also usually mean that no methods will be emitted from this component to the next component in the chain.

Finally, 'cell' method calls corresponding to Empty cells in 'rdsheet' will not be passed on to the next component in the chain.

Calling this component's 'start' method will reset it.

Echo
~~~~

This filter will print calls to the methods configured when the filter is instantiated along with the arguments passed.

MemoryLogger
~~~~~~~~~~~~

This filter will dump stats to the path it was configured with using the heapy package if it is available. If it is not available, no operations are performed.

For more information on heapy, please see `the SourceForge page for heapy <http://guppy-pe.sourceforge.net/#Heapy>_`.

