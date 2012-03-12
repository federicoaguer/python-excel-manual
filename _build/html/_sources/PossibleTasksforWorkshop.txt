Possible Tasks for Workshop
===========================

The following is a list of tasks that can be attempted by any attendee who hasn't brought their own tasks to attempt.

Installation with IronPython
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The libraries have been used successfully with IronPython, but this has not been thoroughly tests or documented.

Installation with Jython
~~~~~~~~~~~~~~~~~~~~~~~~

The libraries should all work with Jython, but no one has so far attempted to do so.

Inserting a row into a sheet
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Starting with an existing Excel file, attempt to create a new Excel file with a row inserted at a given position.

Splitting a Book into its Sheets
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Starting with an existing Excel file, create a directory containing one file for each worksheet in the original file.

Reporting errors in a directory full on Excel files
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Scan a directory of Excel files and report the location of any error cells.

A progression of this task is to allow the passing of options to indicate what types of error to report.

Removing Rows containing errors
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Starting with an existing Excel file, create a filtering process that generates a new Excel file that excludes any rows containing error cells.

A progression of this task is to generate a new Excel file that contains empty cells where there were errors in the original file.

Filtering Excel files to and from a web server
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This task is to create components for xlutils.filter that can read from a website and write back to that website.

The task should result in an HTTPReader and an HTTPWriter.

Producing a report from a database
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This task is to take a typical database query and dump it into an Excel file such that the heading row is set up nicely with decent alignment in a frozen pane.

As a precursor to this task, you may need to set up a typical database!
