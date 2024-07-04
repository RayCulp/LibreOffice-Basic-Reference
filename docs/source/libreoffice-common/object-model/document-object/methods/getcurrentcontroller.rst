Document.GetCurrentController method (LibreOffice common)
=========================================================

Returns the current Controller (the LibreOffice application, for example Calc or Writer) that controls the Document object. The properties and methods available through the Controller object depend on the Controller type.

Syntax
------

*expression*. ``GetCurrentController``

*expression* A variable that represents a Document object. Can also be the ``ThisComponent`` object.

Return value
------------

Object. The Controller object that currently controls the Document object.

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