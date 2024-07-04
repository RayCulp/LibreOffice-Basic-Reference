ThisComponent object (LibreOffice common)
=========================================

Represents either the containing document or the active (foreground) document, depending on where the macro is stored (see Remarks). The properties and methods available through ThisComponent depend on the document type.

Syntax
------

ThisComponent

Examples
--------

.. code-block:: Basic
   
   ' Declarations
   
      Dim oDoc As Object

   ' Get a reference to a the containing or active document into oDoc
   
      oDoc = ThisComponent
   

Remarks
-------

If ThisComponent is used in a macro contained in a specific LibreOffice document, then it refers to the containing document. If ThisComponent is used in a macro that is stored in the **System and User Basic Macro Storage ("My Macros & Dialogs")**, then it refers to the active (foreground) document. 

.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7
