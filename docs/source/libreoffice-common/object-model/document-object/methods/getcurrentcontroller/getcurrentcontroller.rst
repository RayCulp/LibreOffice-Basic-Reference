Document.GetCurrentController method (LibreOffice common)
=========================================================

Get the current controller (i.e. LibreOffice application) that controls the Document object.

Syntax
------

*expression*.GetCurrentController

*expression* A variable that represents a Document object. Can also be the ``ThisComponent`` object.

Remarks
-------

Calling this method is equivalent to choosing the sheet's tab.

Examples
--------

.. code-block:: Basic
   
   ' Declarations
   
      Dim oDoc As Object
      Dim oController As Object

   ' Get a reference to a document into the oDoc variable
   
      oDoc = ThisComponent
   
   ' Get the current controller of oDoc
   
      oController = oDoc.GetCurrentController


.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7