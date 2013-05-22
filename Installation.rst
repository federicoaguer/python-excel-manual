Installation
============

There are several methods of installation available. While the following examples are for
xlrd, the exact same steps can be used for any of the three libraries.

Install from Source
-------------------

On Linux:

::

  $ tar xzf xlrd.tgz
  $ cd xlrd-0.7.1
  $ python setup.py install

**NB:** Make sure you use the python you intend to use for your project.

On Windows, having used WinZip or similar to unpack xlrd-0.7.1.zip:

::

  C:\> cd xlrd-0.7.1
  C:\xlrd-0.7.1> \Python26\python setup.py install


**NB:** Make sure you use the python you intend to use for your project.

Install using Windows Installer
-------------------------------

On Windows, you can download and run the xlrd-0.7.1.win32.exe installer.


Beware that this will only install to Python installations that are in the Windows registry.

Install using EasyInstall
-------------------------

This cross-platform method requires that you already have EasyInstall installed. For more information on this, please see `EasyInstall <http://peak.telecommunity.com/DevCenter/EasyInstall>`_:

::

  easy_install xlrd


Installation using Buildout
---------------------------

Buildout provides a cross-platform method of meeting the python package dependencies of a project without interfering with the system Python.

Having created a directory called ``mybuildout``, download `this file <http://svn.zope.org/*checkout*/zc.buildout/trunk/bootstrap/bootstrap.py>`_ into it. 


Now, create a file in ``mybuildout`` called ``buildout.cfg`` containing the following:

::

  [buildout]
  parts = py 
  versions = versions
  
  [versions]
  xlrd=0.7.1
  xlwt=0.7.2
  xlutils=1.3.2
  
  [py]
  recipe = zc.recipe.egg
  eggs = 
    xlrd 
    xlwt 
    xlutils
  interpreter = py
  
  buildout.cfg

**NB:** The versions section is optional.

Finally, run the following:

::

  $ python bootstrap.py
  $ bin/buildout

These lines:

* initialise the buildout environment

* run the buildout. This should be done each time dependencies change.

Now you can do the following:

::

  $ bin/py your_xlrd_xlwt_xltuils_script.py

Buildout lives at `Buildout <http://pypi.python.org/pypi/zc.buildout>`_.
