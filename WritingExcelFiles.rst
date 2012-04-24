Writing Excel Files
===================

All the examples shown below can be found in the ``xlwt`` directory of the course material.

Creating elements within a Workbook
-----------------------------------

Workbooks are created with ``xlwt`` by instantiating an ``xlwt.Workbook`` object, manipulating it and then calling its ``save`` method.

The ``save`` method may be passed either a string containing the path to write to or a file-like object, opened for writing in binary mode, to which the binary Excel file data will be written.

The following objects can be created within a workbook:

Worksheets
~~~~~~~~~~

Worksheets are created with the ``add_sheet`` method of the ``Workbook`` class.

To retrieve an existing sheet from a ``Workbook``, use its ``get_sheet`` method. This method is particularly useful when the ``Workbook`` has been instantiated by ``xlutils.copy``.

Rows
~~~~

Rows are created using the ``row`` method of the ``Worksheet`` class and contain all of the cells for a given row.

The ``row`` method is also used to retrieve existing rows from a ``Worksheet``.

If a large number of rows have been written to a ``Worksheet`` and memory usage is becoming a problem, the ``flush_row_data`` method may be called on the ``Worksheet``. Once called, any rows flushed cannot be accessed or modified.

It is recommended that ``flush_row_data`` is called for every 1000 or so rows of a normal size that are written to an ``xlwt.Workbook``. If the rows are huge, that number should be reduced.

Columns
~~~~~~~

Columns are created using the ``col`` method of the ``Worksheet`` class and contain display formatting information for a given column.

The ``col`` method is also used to retrieve existing columns from a ``Worksheet``.

Cells
~~~~~

Cells can be written using either the ``write`` method of either the ``Worksheet`` or ``Row`` class.

A more detailed discussion of different ways of writing cells and the different types of cell that may be written is covered later.

A Simple Example
~~~~~~~~~~~~~~~~

The following example shows how all of the above methods can be used to build and save a simple workbook:

::

  from tempfile import TemporaryFile
  from xlwt import Workbook

  book = Workbook()
  sheet1 = book.add_sheet('Sheet 1')
  book.add_sheet('Sheet 2')

  sheet1.write(0,0,'A1')
  sheet1.write(0,1,'B1')
  row1 = sheet1.row(1)
  row1.write(0,'A2')
  row1.write(1,'B2')
  sheet1.col(0).width = 10000

  sheet2 = book.get_sheet(1)
  sheet2.row(0).write(0,'Sheet 2 A1')
  sheet2.row(0).write(1,'Sheet 2 B1')
  sheet2.flush_row_data()
  sheet2.write(1,0,'Sheet 2 A3')
  sheet2.col(0).width = 5000
  sheet2.col(0).hidden = True

  book.save('simple.xls')
  book.save(TemporaryFile())
  
  simple.py

Unicode
--------

The best policy is to pass unicode objects to all ``xlwt``-related method calls.

If you absolutely have to use encoded strings then make sure that the encoding used is consistent across all calls to any ``xlwt``-related methods.

If encoded strings are used and the encoding is not ``'ascii'``, then any ``Workbook`` objects must be created with the appropriate encoding specified:

::

  from xlwt import Workbook
  book = Workbook(encoding='utf-8')

Writing to Cells
----------------

A number of different ways of writing a cell are provided by xlwt along with different strategies for handling multiple writes to the same cell.

Different ways of writing cells
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

There are generally three ways to write to a particular cell:

* ``Worksheet.write(row_index,column_index,value)``

  * This is just syntactic sugar for ``sheet.row(row_index).write(column_index,value)``.

  * It can be useful when you only want to write one cell to a row.

* ``Row.write(column_index,value)``

  * This will write the correct type of cell based on the value passed.

  * Because it figures out what type of cell to write, this method may be slower for writing large workbooks

* Specialist write methods on the Row class

  * Each type of cell has a specialist setter method as covered in the “Types of Cell” section below.

  * These require you to pass the correct type of Python object but can be faster.

In general, use ``Worksheet.write`` for convenience and the specialist write methods if you require speed for a large volume of data.

Overwriting Cells
~~~~~~~~~~~~~~~~~

The Excel file format does nothing to prevent multiple records for a particular cell occurring but, if this happens, the results will vary depending on what application is used to open the file. Excel will display a ``“File error: data may have been lost”`` while OpenOffice.org will show the last record for the cell that occurs in the file.

To help prevent this, ``xlwt`` provides two modes of operation:

* Writing to the same cell more than once will result in an exception.  This is the default mode.

* Writing to the same cell more than once will replace the record for that cell, and only one record will be written when the Workbook is saved.

The following example demonstrates these two options:

::

  from xlwt import Workbook

  book = Workbook()
  sheet1 = book.add_sheet('Sheet 1',cell_overwrite_ok=True)
  sheet1.write(0,0,'original')
  sheet = book.get_sheet(0)
  sheet.write(0,0,'new')

  sheet2 = book.add_sheet('Sheet 2')
  sheet2.write(0,0,'original')
  sheet2.write(0,0,'new')
  
  overwriting.py

The most common case for needing to overwrite cells is when an existing Excel file has been loaded into a Workbook instance using ``xlutils.copy``.

Types of Cell
-------------

All types of cell supported by the Excel file format can be written:

Text
~~~~

When passed a ``unicode`` or string, the ``write`` methods will write a ``Text`` cell.

The ``set_cell_text`` method of the ``Row`` class can also be used to write ``Text`` cells.

When passed a string, these methods will first decode the string using the Workbook's encoding.

Number
~~~~~~

When passed a ``float``, ``int``, ``long``, or ``decimal.Decimal``, the ``write`` methods will write a ``Number`` cell.

The ``set_cell_number`` method of the ``Row`` class can also be used to write ``Number`` cells.

Date
~~~~

When passed a ``datetime.datetime``, ``datetime.date``, or ``datetime.time``, the ``write`` methods will write a ``Date`` cell.

The ``set_cell_date`` method of the ``Row`` class can also be used to write ``Date`` cells.

**Note**: As mentioned earlier, a date is not really a separate type in Excel; if you don't apply a date format, it will be treated as a number.

Boolean
~~~~~~~

When passed a ``bool``, the ``write`` methods will write a ``Boolean`` cell.

The ``set_cell_boolean`` method of the ``Row`` class can also be used to write ``Text`` cells.

Error
~~~~~

You shouldn't ever want to write ``Error`` cells!

However, if you absolutely must, the ``set_cell_error`` method of the Row class can be used to do so. For convenience, it can be called with either hexadecimal error codes, expressed as integers, or the error text that Excel would display.

Blank
~~~~~

It is not normally necessary to write blank cells. The one exception to this is if you wish to apply formatting to a cell that contains nothing.

To do this, either call the ``write`` methods with an empty string or ``None``, or use the ``set_cell_blank`` method of the ``Row`` class.

If you need to do this for more than one cell in a row, using the ``set_cell_mulblanks`` method will result in a smaller Excel file when the ``Workbook`` is saved.

The following example brings all of the above cell types together and shows examples use both the generic ``write`` method and the specialist methods:

::

  from datetime import date,time,datetime
  from decimal import Decimal
  from xlwt import Workbook,Style

  wb = Workbook()
  ws = wb.add_sheet('Type examples')
  ws.row(0).write(0,u'\xa3')
  ws.row(0).write(1,'Text')
  ws.row(1).write(0,3.1415)
  ws.row(1).write(1,15)
  ws.row(1).write(2,265L)
  ws.row(1).write(3,Decimal('3.65'))
  ws.row(2).set_cell_number(0,3.1415)
  ws.row(2).set_cell_number(1,15)
  ws.row(2).set_cell_number(2,265L)
  ws.row(2).set_cell_number(3,Decimal('3.65'))
  ws.row(3).write(0,date(2009,3,18))
  ws.row(3).write(1,datetime(2009,3,18,17,0,1))
  ws.row(3).write(2,time(17,1))
  ws.row(4).set_cell_date(0,date(2009,3,18))
  ws.row(4).set_cell_date(1,datetime(2009,3,18,17,0,1))
  ws.row(4).set_cell_date(2,time(17,1))
  ws.row(5).write(0,False)
  ws.row(5).write(1,True)
  ws.row(6).set_cell_boolean(0,False)
  ws.row(6).set_cell_boolean(1,True)
  ws.row(7).set_cell_error(0,0x17)
  ws.row(7).set_cell_error(1,'#NULL!')
  ws.row(8).write(
      0,'',Style.easyxf('pattern: pattern solid, fore_colour green;'))
  ws.row(8).write(
      1,None,Style.easyxf('pattern: pattern solid, fore_colour blue;'))
  ws.row(9).set_cell_blank(
      0,Style.easyxf('pattern: pattern solid, fore_colour yellow;'))
  ws.row(10).set_cell_mulblanks(
      5,10,Style.easyxf('pattern: pattern solid, fore_colour red;')
      )

  wb.save('types.xls')

  cell_types.py

Styles
------

Most elements of an Excel file can be formatted. For many elements including cells, rows and columns, this is done by assigning a style, known as an XF record, to that element.

This is done by passing an ``xlwt.XFStyle`` instance to the optional last argument to the various write methods and specialist ``set_cell_ methods``. ``xlwt.Row`` and ``xlwt.Column`` instances have ``set_style`` methods to which an ``xlwt.XFStyle`` instance can be passed.

XFStyle
~~~~~~~

In ``xlwt``, the XF record is represented by the ``XFStyle`` class and its related attribute classes.

The following example shows how to create a red ``Date`` cell with Arial text and a black border:

::

  from datetime import date
  from xlwt import Workbook, XFStyle, Borders, Pattern, Font

  fnt = Font()
  fnt.name = 'Arial'

  borders = Borders()
  borders.left = Borders.THICK
  borders.right = Borders.THICK
  borders.top = Borders.THICK
  borders.bottom = Borders.THICK

  pattern = Pattern()
  pattern.pattern = Pattern.SOLID_PATTERN
  pattern.pattern_fore_colour = 0x0A

  style = XFStyle()
  style.num_format_str='YYYY-MM-DD'
  style.font = fnt
  style.borders = borders
  style.pattern = pattern

  book = Workbook()
  sheet = book.add_sheet('A Date')
  sheet.write(1,1,date(2009,3,18),style)

  book.save('date.xls')

  xfstyle_format.py

This can be quite cumbersome!

easyxf
~~~~~~

Thankfully, ``xlwt`` provides the ``easyxf`` helper to create ``XFStyle`` instances from human readable text and an optional string containing a number format.

Here is the above example, this time created with ``easyxf``:

::

  from datetime import date
  from xlwt import Workbook, easyxf

  book = Workbook()
  sheet = book.add_sheet('A Date')

  sheet.write(1,1,date(2009,3,18),easyxf(
      'font: name Arial;'
      'borders: left thick, right thick, top thick, bottom thick;'
      'pattern: pattern solid, fore_colour red;',
      num_format_str='YYYY-MM-DD'
      ))

  book.save('date.xls')

  easyxf_format.py

The human readable text breaks roughly as follows, in pseudo-regular expression syntax:

``(<element>:(<attribute> <value>,)+;)+``

This means:

* The text contains a semi-colon delimited list of element definitions.

* Each element contains a comma-delimited list of attribute and value pairs.

The following sections describe each of the types of element by providing a table of their attributes and possible values for those attributes. For explanations of how to express boolean values and colours, please see the “Types of attribute” section.

font
~~~~

+-----------------+---------------------------------------------------------------------------------------+
| ``bold``        | A *boolean* value.                                                                    |
|                 | The default is ``False``.                                                             |        
+-----------------+---------------------------------------------------------------------------------------+
| ``charset``     | The character set to use for this font, which can be one of the following:            |
|                 | ``ansi_latin``, ``sys_default``, ``symbol``, ``apple_roman``, ``ansi_jap_shift_jis``, |
|                 | ``ansi_kor_hangul``, ``ansi_kor_johab``, ``ansi_chinese_gbk``, ``ansi_chinese_big5``, |
|                 | ``ansi_greek``, ``ansi_turkish``, ``ansi_vietnamese``, ``ansi_hebrew``,               |
|                 | ``ansi_arabic``, ``ansi_baltic``, ``ansi_cyrillic``, ``ansi_thai``, ``ansi_latin_ii``,|
|                 | ``oem_latin_i``.                                                                      |
|                 | The default is ``sys_default``.                                                       |
+-----------------+---------------------------------------------------------------------------------------+
| ``colour``      | A *colour* specifying the colour for the text.                                        |
|                 | The default is the automatic colour.                                                  |
+-----------------+---------------------------------------------------------------------------------------+
| ``escapement``  | This can be one of ``none``, ``superscript``, or ``subscript``.                       |
|                 | The default is  ``none``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``family``      | This should be a string containing the name of the font family to use. You probably   |
|                 | want to use ``name`` instead of this attribute and leave this to its default value.   |
|                 | The default is ``None``.                                                              |
+-----------------+---------------------------------------------------------------------------------------+
| ``height``      | The height of the font as expressed by multiplying the point size by 20.              |
|                 | The default is 200, which equates to 10pt.                                            |
+-----------------+---------------------------------------------------------------------------------------+
| ``italic``      | A *boolean* value.                                                                    |
|                 | The default is ``False``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``name``        | This should be a string containing the name of the font family to use.                |
|                 | The default is ``Arial``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``outline``     | A *boolean*  value.                                                                   |
|                 | The default is ``False``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``shadow``      | A *boolean* value.                                                                    |
|                 | The default is ``False``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``struck_out``  | A *boolean* value.                                                                    |
|                 | The default is ``False``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``underline``   | A *boolean* value or one of  ``none``, ``single``, ``single_acc``, ``double``, or     |
|                 | ``double_acc``.                                                                       |
|                 | The default is ``none``.                                                              |
+-----------------+---------------------------------------------------------------------------------------+
| ``color_index`` | A synonym for ``colour``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``colour_index``| A synonym for ``colour``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+
| ``color``       | A synonym for ``colour``.                                                             |
+-----------------+---------------------------------------------------------------------------------------+

alignment
~~~~~~~~~

+-------------------+-----------------------------------------------------------------------------------------------------+
| ``direction``     | One of ``general``, ``lr``, or ``rl``.                                                              |
|                   | The default is ``general``.                                                                         |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``horizontal``    | One of the following: ``general``, ``left``, ``center|centre``, ``right``, ``filled``, ``justified``|
|                   | , ``center|centre_across_selection``, ``distributed``.                                              |
|                   | The default is ``general``.                                                                         |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``indent``        | A indentation amount between 0 and 15.                                                              |
|                   | The default is 0.                                                                                   |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``rotation``      | An integer rotation in degrees between -90 and +90 or one of ``stacked`` or ``none``.               |
|                   | The default is ``none``.                                                                            |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``shrink_to_fit`` | A *boolean* value.                                                                                  |
|                   | The default is ``False``.                                                                           |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``vertical``      | One of the following: ``top``, ``center``/``centre``, ``bottom``, ``justified``, ``distributed``.   |
|                   | The default is ``bottom``.                                                                          |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``wrap``          | A *boolean* value.                                                                                  |
|                   | The default is ``False``.                                                                           |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``dire``          | This is a synonym for ``direction``.                                                                |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``horiz``         | This is a synonym for ``horizontal``.                                                               |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``horz``          | This is a synonym for ``horizontal``.                                                               |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``inde``          | This is a synonym for ``indent``.                                                                   |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``rota``          | This is a synonym for ``rotation``.                                                                 |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``shri``          | This is a synonym for ``shrink_to_fit``.                                                            |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``shrink``        | This is a synonym for ``shrink_to_fit``.                                                            |
+-------------------+-----------------------------------------------------------------------------------------------------+
| ``vert``          | This is a synonym for ``vertical``.                                                                 |
+-------------------+-----------------------------------------------------------------------------------------------------+

borders
~~~~~~~

+-------------------+------------------------------------------+
| ``left``          | A type of border line.*                  |
+-------------------+------------------------------------------+
| ``right``         | A type of border line.*                  |
+-------------------+------------------------------------------+
| ``top``           | A type of border line.*                  |
+-------------------+------------------------------------------+
| ``bottom``        | A type of border line.*                  |
+-------------------+------------------------------------------+
| ``diag``          | A type of border line.*                  |
+-------------------+------------------------------------------+
| ``left_colour``   | A *colour*.                              |
|                   | The default is the ``automatic`` colour. |
+-------------------+------------------------------------------+
| ``right_colour``  | A *colour*.                              |
|                   | The default is the ``automatic`` colour. |
+-------------------+------------------------------------------+
| ``top_colour``    | A *colour*.                              |
|                   | The default is the ``automatic`` colour. |
+-------------------+------------------------------------------+
| ``bottom_colour`` | A *colour*.                              |
|                   | The default is the ``automatic`` colour. |
+-------------------+------------------------------------------+
| ``diag_colour``   | A *colour*.                              |
|                   | The default is the ``automatic`` colour. |
+-------------------+------------------------------------------+
| ``need_diag_1``   | A *boolean* value.                       |
|                   | The default is ``False``.                |
+-------------------+------------------------------------------+
| ``need_diag_2``   | A *boolean* value.                       |
|                   | The default is ``False``.                |
+-------------------+------------------------------------------+
| ``left_color``    | A synonym for ``left_colour``.           |
+-------------------+------------------------------------------+
| ``right_color``   | A synonym for ``right_colour``.          |
+-------------------+------------------------------------------+
| ``top_color``     | A synonym for ``top_colour``.            |
+-------------------+------------------------------------------+
| ``bottom_color``  | A synonym for ``bottom_colour``.         |
+-------------------+------------------------------------------+
| ``diag_color``    | A synonym for ``diag_colour``.           |
+-------------------+------------------------------------------+

* This can be either an integer width between 0 and 13 or one of the following: ``no_line``, ``thin``, ``medium``, ``dashed``, ``dotted``, ``thick``, ``double``, ``hair``, ``medium_dashed``, ``thin_dash_dotted``, ``medium_dash_dotted``, ``thin_dash_dot_dotted``, ``medium_dash_dot_dotted``, or ``slanted_medium_dash_dotted``.

pattern
~~~~~~~

+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``back_colour``         | A *colour*.                                                                                              |
|                         | The default is the ``automatic`` colour.                                                                 |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``fore_colour``         | A *colour*.                                                                                              |
|                         | The default is  the ``automatic`` colour.                                                                |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``pattern``             | One of the following: ``no_fill``, ``none``, ``solid``, ``solid_fill``, ``solid_pattern``, ``fine_dots``,|
|                         | ``alt_bars``, ``sparse_dots``, ``thick_horz_bands``, ``thick_vert_bands``, ``thick_backward_diag``,      |
|                         | ``thick_forward_diag``, ``big_spots``, ``bricks``, ``thin_horz_bands``, ``thin_vert_bands``,             |
|                         | ``thin_backward_diag``, ``thin_forward_diag``, ``squares``, or ``diamonds``.                             |
|                         | The default is ``none``.                                                                                 |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``fore_color``          | A synonym  for ``fore_colour``.                                                                          |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``back_color``          | A synonym for ``back_colour``.                                                                           |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``pattern_fore_colour`` | A synonym for ``fore_colour``.                                                                           |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``pattern_fore_color``  | A synonym for ``fore_colour``.                                                                           |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``pattern_back_colour`` | A synonym for ``back_colour``.                                                                           |
+-------------------------+----------------------------------------------------------------------------------------------------------+
| ``pattern_back_color``  | A synonym for ``back_colour``.                                                                           |
+-------------------------+----------------------------------------------------------------------------------------------------------+

protection
~~~~~~~~~~

The protection features of the Excel file format are only partially implemented in ``xlwt``. Avoid them unless you plan on finishing their implementation.

+--------------------+--------------------------+
| ``cell_locked``    | A *boolean* value.       |
|                    | The default is ``True``. |
+--------------------+--------------------------+
| ``formula_hidden`` | A *boolean* value.       |
|                    | The default is ``False``.|
+--------------------+--------------------------+

align
~~~~~

A synonym for ``alignment``.

border
~~~~~~

A synonym for ``borders``

Types of attributes
~~~~~~~~~~~~~~~~~~~

*Boolean* values are either ``True`` or ``False``, but ``easyxf`` allows great flexibility in how you choose to express those two values:

* ``True`` can be expressed by ``1``, ``yes``, ``true``, or ``on``.

* ``False`` can be expressed by ``0``, ``no``, ``false``, or ``off``.

*Colours* in Excel files are a confusing mess. The safest bet to do is just pick from the following list of colour names that ``easyxf`` understands.

The names used are those reported by the Excel 2003 GUI when you are inspecting the **default** colour palette.

Warning: There are many differences between this implicit mapping from colour-names to RGB values and the mapping used in standards such as HTML and CSS.

+--------------------+------------------+---------------------+----------------+
| ``aqua``           | ``dark_red_ega`` | ``light_blue``      | ``plum``       |
+--------------------+------------------+---------------------+----------------+
| ``black``          | ``dark_teal``    | ``light_green``     | ``purple_ega`` |
+--------------------+------------------+---------------------+----------------+
| ``blue``           | ``dark_yellow``  | ``light_orange``    | ``red``        |
+--------------------+------------------+---------------------+----------------+
| ``blue_gray``      | ``gold``         | ``light_turquoise`` | ``rose``       |
+--------------------+------------------+---------------------+----------------+
| ``bright_green``   | ``gray_ega``     | ``light_yellow``    | ``sea_green``  |
+--------------------+------------------+---------------------+----------------+
| ``brown``          | ``gray25``       | ``lime``            | ``silver_ega`` |
+--------------------+------------------+---------------------+----------------+
| ``coral``          | ``gray40``       | ``magenta_ega``     | ``sky_blue``   |
+--------------------+------------------+---------------------+----------------+
| ``cyan_ega``       | ``gray50``       | ``ocean_blue``      | ``tan``        |
+--------------------+------------------+---------------------+----------------+
| ``dark_blue``      | ``gray80``       | ``olive_ega``       | ``teal``       |
+--------------------+------------------+---------------------+----------------+
| ``dark_blue_ega``  | ``green``        | ``olive_green``     | ``teal_ega``   |
+--------------------+------------------+---------------------+----------------+
| ``dark_green``     | ``ice_blue``     | ``orange``          | ``turquoise``  |
+--------------------+------------------+---------------------+----------------+
| ``dark_green_ega`` | ``indigo``       | ``pale_blue``       | ``violet``     |
+--------------------+------------------+---------------------+----------------+
| ``dark_purple``    | ``ivory``        | ``periwinkle``      | ``white``      |
+--------------------+------------------+---------------------+----------------+
| ``dark_red``       | ``lavender``     | ``pink``            | ``yellow``     |
+--------------------+------------------+---------------------+----------------+

**NB**: ``grey`` can be used instead of ``gray`` wherever it occurs above.

Formatting Rows and Columns
---------------------------

It is possible to specify default formatting for rows and columns within a worksheet. This is done using the ``set_style`` method of the ``Row`` and ``Column`` instances, respectively.

The precedence of styles is as follows:

* the style applied to a cell

* the style applied to a row

* the style applied to a column

It is also possible to hide whole rows and columns by using the ``hidden`` attribute of ``Row`` and ``Column`` instances.

The width of a ``Column`` can be controlled by setting its ``width`` attribute to an integer where 1 is 1/256 of the width of the zero character, using the first font that occurs in the Excel file.

By default, the height of a row is determined by the tallest font for that row and the ``height`` attribute of the row is ignored.
If you want the ``height`` attribute to be used, the row's ``height_mismatch`` attribute needs to be set to ``1``.

The following example shows these methods and properties in use along with the style precedence:

::
  
  from xlwt import Workbook, easyxf
  from xlwt.Utils import rowcol_to_cell
  
  row = easyxf('pattern: pattern solid, fore_colour blue')
  col = easyxf('pattern: pattern solid, fore_colour green')
  cell = easyxf('pattern: pattern solid, fore_colour red')
  
  book = Workbook()
  
  sheet = book.add_sheet('Precedence')
  for i in range(0,10,2):
      sheet.row(i).set_style(row)
  for i in range(0,10,2):
      sheet.col(i).set_style(col)
  for i in range(10):
      sheet.write(i,i,None,cell)
  
  sheet = book.add_sheet('Hiding')
  for rowx in range(10):
      for colx in range(10):
          sheet.write(rowx,colx,rowcol_to_cell(rowx,colx))                    
  for i in range(0,10,2):
      sheet.row(i).hidden = True
      sheet.col(i).hidden = True
  
  sheet = book.add_sheet('Row height and Column width')
  for i in range(10):
      sheet.write(0,i,0)
  for i in range(10):
      sheet.row(i).set_style(easyxf('font:height '+str(200*i)))
      sheet.col(i).width = 256*i
  
  book.save('format_rowscols.xls')
  format_rowscols.py

Formatting Sheets and Workbooks
-------------------------------

There are many possible settings that can be made on Sheets and Workbooks.

Most of them you will never need or want to touch.

If you think you do, see the “Other Properties” section below.

Style compression
-----------------

While it is fine to create as many XFStyles and their associated Font instances as you like, each one written to ``Workbook`` will result in an XF record and a Font record. Excel has fixed limits of around 400 Fonts and 4000 XF records so care needs to be taken when generating large Excel files.

To help with this, ``xlwt.Workbook`` has an optional ``style_compression`` parameter with the following meaning:

* 0 – no compression. This is the default.

* 1 – compress Fonts only. Not very useful.

* 2 – compress Fonts and XF records.

The following example demonstrates these three options:

::

  from xlwt import Workbook, easyxf
  
  style1 = easyxf('font: name Times New Roman')
  style2 = easyxf('font: name Times New Roman')
  style3 = easyxf('font: name Times New Roman')
  
  def write_cells(book):
      sheet = book.add_sheet('Content')
      sheet.write(0,0,'A1',style1)
      sheet.write(0,1,'B1',style2)
      sheet.write(0,2,'C1',style3)
      
  book = Workbook()
  write_cells(book)
  book.save('3xf3fonts.xls')
  
  book = Workbook(style_compression=1)
  write_cells(book)
  book.save('3xf1font.xls')
  
  book = Workbook(style_compression=2)
  write_cells(book)
  book.save('1xf1font.xls')
  stylecompression.py

Be aware that doing this compression involves deeply nested comparison of the XFStyle objects, so may slow down writing of large files where many styles are used.

The recommended best practice is to create all the styles you will need in advance and leave ``style_compression`` at its default value.

Formulae
--------

Formulae can be written by ``xlwt`` by passing an ``xlwt.Formula`` instance to either of the write methods or by using the ``set_cell_formula`` method of ``Row`` instances, bugs allowing.

The following are supported:

* all the built-in Excel formula functions

* references to other sheets in the same workbook

* access to all the add-in functions in the Analysis Toolpak (ATP)

* comma or semicolon as the argument separator in function calls

* case-insensitive matching of formula names

The following are not suppoted:

* references to external workbooks

* array aka Ctrl-Shift-Enter aka CSE formulas

* references to defined Names

* using formulas for data validation or conditional formatting

* evaluation of formulae

The following example shows some of these things in action:

::

  from xlwt import Workbook, Formula
  
  book = Workbook()
  
  sheet1 = book.add_sheet('Sheet 1')
  sheet1.write(0,0,10)
  sheet1.write(0,1,20)
  sheet1.write(1,0,Formula('A1/B1'))
  
  sheet2 = book.add_sheet('Sheet 2')
  row = sheet2.row(0)
  row.write(0,Formula('sum(1,2,3)'))
  row.write(1,Formula('SuM(1;2;3)'))
  row.write(2,Formula("$A$1+$B$1*SUM('ShEEt 1'!$A$1:$b$2)"))
  
  book.save('formula.xls')
  formulae.py

Names
-----

Names cannot currently be written by ``xlwt``.

Utility methods
---------------

The ``Utils`` module of ``xlwt`` contains several useful utility functions:

col_by_name
~~~~~~~~~~~

This will convert a string containing a column identifier into an integer column index.

cell_to_rowcol
~~~~~~~~~~~~~~

This will convert a string containing an excel cell reference into a four-element tuple containing: 

``(row,col,row_abs,col_abs)``

``row``
– integer row index of the referenced cell

``col``
– integer column index of the referenced cell

``row_abs``
– *boolean* indicating whether the row index is absolute (``True``) or relative (``False``)

``col_abs``
– *boolean* indicating whether the column index is absolute (``True``) or relative (``False``)

cell_to_rowcol2
~~~~~~~~~~~~~~~

This will convert a string containing an excel cell reference into a two-element tuple containing:

``(row,col)``

``row``
– integer row index of the referenced cell

``col``
– integer column index of the referenced cell

rowcol_to_cell
~~~~~~~~~~~~~~

This will covert an integer row and column index into a string excel cell reference, with either index optionally being absolute.

cellrange_to_rowcol_pair
~~~~~~~~~~~~~~~~~~~~~~~~

This will convert a string containing an excel range into a four-element tuple containing:

``(row1,col1,row2,col2)``

``row1``
– integer row index of the start of the range

``col1``
– integer column index of the start of the range

``row2``
– integer row index of the end of the range

``col2``
– integer column index of the end of the range

rowcol_pair_to_cellrange
~~~~~~~~~~~~~~~~~~~~~~~~

This will covert a pair of integer row and column indexes into a string containing an Excel cell range. Any of the
indexes specified can optionally be made to be absolute.

valid_sheet_name
~~~~~~~~~~~~~~~~

This function takes a single string argument and returns a *boolean* value indication whether the sheet name will work without problems (``True``) or will cause complaints from Excel (``False``).

The following example shows all of these functions in use:

::

  from xlwt import Utils
  
  print 'AA ->',Utils.col_by_name('AA')
  print 'A ->',Utils.col_by_name('A')
  
  print 'A1 ->',Utils.cell_to_rowcol('A1')
  print '$A$1 ->',Utils.cell_to_rowcol('$A$1')
  
  print 'A1 ->',Utils.cell_to_rowcol2('A1')
  
  print (0,0),'->',Utils.rowcol_to_cell(0,0)
  print (0,0,False,True),'->',
  print Utils.rowcol_to_cell(0,0,False,True)
  print (0,0,True,True),'->',
  print Utils.rowcol_to_cell(
            row=0,col=0,row_abs=True,col_abs=True
            )
  
  print '1:3 ->',Utils.cellrange_to_rowcol_pair('1:3')
  print 'B:G ->',Utils.cellrange_to_rowcol_pair('B:G')
  print 'A2:B7 ->',Utils.cellrange_to_rowcol_pair('A2:B7')
  print 'A1 ->',Utils.cellrange_to_rowcol_pair('A1')
  
  print (0,0,100,100),'->',
  print Utils.rowcol_pair_to_cellrange(0,0,100,100)
  print (0,0,100,100,True,False,False,False),'->',
  print Utils.rowcol_pair_to_cellrange(
            row1=0,col1=0,row2=100,col2=100,
            row1_abs=True,col1_abs=False,
            row2_abs=False,col2_abs=True
            )
  
  for name in (
      '',"'quoted'","O'hare","X"*32,"[]:\\?/*\x00"
      ):
      print 'Is %r a valid sheet name?' % name,
      if Utils.valid_sheet_name(name):
          print "Yes"
      else:
          print "No"
  utilities.py

Other properties
----------------

There are many other properties that you can set on ``xlwt``-related objects. They are all listed below, for each of the types of object. The names are mostly intuitive but you are warned to experiment thoroughly before attempting to use any of these in an important situation as some properties exist that aren't saved to the resulting Excel files and some others are only partially implemented.

xlwt.Workbook
~~~~~~~~~~~~~

+--------------------+------------------+---------------------+
| ``owner``          | ``vpos``         | ``hscroll_visible`` |
+--------------------+------------------+---------------------+
| ``country_code``   | ``width``        | ``vscroll_visible`` |
+--------------------+------------------+---------------------+
| ``wnd_protect``    | ``height``       | ``tabs_visible``    |
+--------------------+------------------+---------------------+
| ``obj_protect``    | ``active_sheet`` | ``dates_1904``      |
+--------------------+------------------+---------------------+
| ``protect``        | ``tab_width``    | ``use_cell_values`` |
+--------------------+------------------+---------------------+
| ``backup_on_save`` | ``wnd_visible``  |                     |
+--------------------+------------------+---------------------+
| ``hpos``           | ``wnd_mini``     |                     |
+--------------------+------------------+---------------------+

xlwt.Row
~~~~~~~~

+------------------------+---------------------+-----------------+
| ``set_style``          | ``height_mismatch`` | ``hidden``      |
+------------------------+---------------------+-----------------+
| ``height``             | ``level``           | ``space_above`` |
+------------------------+---------------------+-----------------+
| ``has_default_height`` | ``collapse``        | ``space_below`` |
+------------------------+---------------------+-----------------+

xlwt.Column
~~~~~~~~~~~

+---------------------+------------+--------------+
| ``set_style``       | ``width``  | ``level``    |
| ``width_in_pixels`` | ``hidden`` | ``collapse`` |
+---------------------+------------+--------------+

xlwt.Worksheet
~~~~~~~~~~~~~~

+---------------------------------+-------------------------+
| ``name``                        | ``save_recalc``         |
+---------------------------------+-------------------------+
| ``visibility``                  | ``print_headers``       |
+---------------------------------+-------------------------+
| ``row_default_height_mismatch`` | ``print_grid``          |
+---------------------------------+-------------------------+
| ``row_default_hidden``          | ``header_str``          |
+---------------------------------+-------------------------+
| ``row_default_space_above``     | ``footer_str``          |
+---------------------------------+-------------------------+
| ``row_default_space_below``     | ``print_centered_vert`` |
+---------------------------------+-------------------------+
| ``show_formulas``               | ``print_centered_horz`` |
+---------------------------------+-------------------------+
| ``show_grid``                   | ``left_margin``         |
+---------------------------------+-------------------------+
| ``show_headers``                | ``right_margin``        |
+---------------------------------+-------------------------+
| ``show_zero_values``            | ``top_margin``          |
+---------------------------------+-------------------------+
| ``auto_colour_grid``            | ``bottom_margin``       |
+---------------------------------+-------------------------+
| ``cols_right_to_left``          | ``paper_size_code``     |
+---------------------------------+-------------------------+
| ``show_outline``                | ``print_scaling``       |
+---------------------------------+-------------------------+
| ``remove_splits``               | ``start_page_number``   |
+---------------------------------+-------------------------+
| ``selected``                    | ``fit_width_to_pages``  |
+---------------------------------+-------------------------+
| ``sheet_visible``               | ``fit_height_to_pages`` |
+---------------------------------+-------------------------+
| ``page_preview``                | ``print_in_rows``       |
+---------------------------------+-------------------------+
| ``first_visible_row``           | ``portrait``            |
+---------------------------------+-------------------------+
| ``first_visible_col``           | ``print_colour``        |
+---------------------------------+-------------------------+
| ``grid_colour``                 | ``print_draft``         |
+---------------------------------+-------------------------+
| ``dialog_sheet``                | ``print_notes``         |
+---------------------------------+-------------------------+
| ``auto_style_outline``          | ``print_notes_at_end``  |
+---------------------------------+-------------------------+
| ``outline_below``               | ``print_omit_errors``   |
+---------------------------------+-------------------------+
| ``outline_right``               | ``print_hres``          |
+---------------------------------+-------------------------+
| ``fit_num_pages``               | ``header_margin``       |
+---------------------------------+-------------------------+
| ``show_row_outline``            | ``footer_margin``       |
+---------------------------------+-------------------------+
| ``show_col_outline``            | ``copies_num``          |
+---------------------------------+-------------------------+
| ``alt_expr_eval``               | ``wnd_protect``         |
+---------------------------------+-------------------------+
| ``alt_formula_entries``         | ``obj_protect``         |
+---------------------------------+-------------------------+
| ``row_default_height``          | ``protect``             |
+---------------------------------+-------------------------+
| ``col_default_height``          | ``scen_protect``        |
+---------------------------------+-------------------------+
| ``calc_mode``                   | ``password``            |
+---------------------------------+-------------------------+
| ``calc_count``                  |                         |
+---------------------------------+-------------------------+
| ``RC_ref_mode``                 |                         |
+---------------------------------+-------------------------+
| ``iterations_on``               |                         |
+---------------------------------+-------------------------+
| ``delta``                       |                         |
+---------------------------------+-------------------------+

Some examples of Other Properties
---------------------------------

The following sections contain examples of how to use some of the properties listed above.

Hyperlinks
~~~~~~~~~~

Hyperlinks are a type of formula as shown in the following example:

::

  from xlwt import Workbook,easyxf,Formula
  
  style = easyxf('font: underline single')
  
  book = Workbook()
  sheet = book.add_sheet('Hyperlinks')
  
  sheet.write(
      0, 0,
      Formula('HYPERLINK("http://www.python.org";"Python")'),
      style)
  
  link = 'HYPERLINK("mailto:python-excel@googlegroups.com";"help")'
  sheet.write(
      1,0,
      Formula(link),
      style)
  
  book.save("hyperlinks.xls")
  hyperlinks.py

Images
~~~~~~~

Images can be inserted using the ``insert_bitmap`` method of the ``Sheet`` class:

::

  from xlwt import Workbook
  w = Workbook()
  ws = w.add_sheet('Image')
  ws.insert_bitmap('python.bmp', 0, 0)
  w.save('images.xls')
  images.py

**NB**: Images are not displayed by ``OpenOffice.org``.

Merged cells
~~~~~~~~~~~~

Merged groups of cells can be inserted using the ``write_merge`` method of the ``Sheet`` class:

::

  from xlwt import Workbook,easyxf
  style = easyxf(
      'pattern: pattern solid, fore_colour red;'
      'align: vertical center, horizontal center;'
      )
  w = Workbook()
  ws = w.add_sheet('Merged')
  ws.write_merge(1,5,1,5,'Merged',style)
  w.save('merged.xls')
  merged.py

Borders
~~~~~~~

Writing a single cell with borders is simple enough, however applying a border to a group of cells is painful as shown in this example:

::

  from xlwt import Workbook,easyxf
  tl = easyxf('border: left thick, top thick')
  t = easyxf('border: top thick')
  tr = easyxf('border: right thick, top thick')
  r = easyxf('border: right thick')
  br = easyxf('border: right thick, bottom thick')
  b = easyxf('border: bottom thick')
  bl = easyxf('border: left thick, bottom thick')
  l = easyxf('border: left thick')
  
  w = Workbook()
  ws = w.add_sheet('Border')
  ws.write(1,1,style=tl)
  ws.write(1,2,style=t)
  ws.write(1,3,style=tr)
  ws.write(2,3,style=r)
  ws.write(3,3,style=br)
  ws.write(3,2,style=b)
  ws.write(3,1,style=bl)
  ws.write(2,1,style=l)
  
  w.save('borders.xls')
  borders.py

**NB**: Extra care needs to be taken if you're updating an existing Excel file!

Split and Freeze panes
~~~~~~~~~~~~~~~~~~~~~~

It is fairly straight forward to create frozen panes using ``xlwt``.

The location of the split is specified using the integer ``vert_split_pos`` and ``horz_split_pos`` properties of the ``Sheet`` class.

The first visible cells are specified using the integer ``vert_split_first_visible`` and ``horz_split_first_visible`` properties of the ``Sheet`` class.

The following example shows them all in action:

::

  from xlwt import Workbook
  from xlwt.Utils import rowcol_to_cell
  
  w = Workbook()
  sheet = w.add_sheet('Freeze')
  sheet.panes_frozen = True
  sheet.remove_splits = True
  sheet.vert_split_pos = 2
  sheet.horz_split_pos = 10
  sheet.vert_split_first_visible = 5
  sheet.horz_split_first_visible = 40
  
  for col in range(20):
      for row in range(80):
          sheet.write(row,col,rowcol_to_cell(row,col))
  
  w.save('panes.xls')
  panes.py

Split panes are a less frequently used feature and their support is less complete in ``xlwt``.

The procedure for creating split panes is exactly the same as for frozen panes except that the ``panes_frozen`` attribute of the Worksheet should be set to ``False`` instead of ``True``.

However, if you really need split panes, you're advised to see professional help before proceeding!

Outlines
~~~~~~~~~

These are a little known and little used feature of the Excel file format that can be very useful when dealing with categorised data.

Their use is best shown by example:

::

  from xlwt import Workbook
  data = [
      ['','','2008','','2009'],
      ['','','Jan','Feb','Jan','Feb'],
      ['Company X'],
      ['','Division A'],
      ['','',100,200,300,400],
      ['','Division B'],
      ['','',100,99,98,50],
      ['Company Y'],
      ['','Division A'],
      ['','',100,100,100,100],
      ['','Division B'],
      ['','',100,101,102,103],
      ]
  w = Workbook()
  ws = w.add_sheet('Outlines')
  for i,row in enumerate(data):
      for j,cell in enumerate(row):
          ws.write(i,j,cell)
  ws.row(2).level = 1
  ws.row(3).level = 2
  ws.row(4).level = 3
  ws.row(5).level = 2
  ws.row(6).level = 3
  ws.row(7).level = 1
  ws.row(8).level = 2
  ws.row(9).level = 3
  ws.row(10).level = 2
  ws.row(11).level = 3
  ws.col(2).level = 1
  ws.col(3).level = 2
  ws.col(4).level = 1
  ws.col(5).level = 2
  w.save('outlines.xls')
  outlines.py


Zoom magnification and Page Break Preview
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The zoom percentage used when viewing a sheet in ``normal`` mode can be controlled by setting the ``normal_magn`` attribute of a ``Sheet`` instance.

The zoom percentage used when viewing a sheet in ``page break preview`` mode can be controlled by setting the ``preview_magn`` attribute of a ``Sheet`` instance.

A ``Sheet`` can also be made to show a ``page break preview`` by setting the ``page_preview`` attribute of the ``Sheet`` instance to ``True``.

Here's an example to show all three in action:

::

  from xlwt import Workbook
  
  w = Workbook()
  
  ws = w.add_sheet('Normal')
  ws.write(0,0,'Some text')
  ws.normal_magn = 75
  
  ws = w.add_sheet('Page Break Preview')
  ws.write(0,0,'Some text')
  ws.preview_magn = 150
  ws.page_preview = True
  
  w.save('zoom.xls')
  zoom.py
