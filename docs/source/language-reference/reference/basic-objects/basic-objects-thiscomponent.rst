ThisComponent object
====================

The object **ThisComponent** is a predefined global variable that provides access to the current component (or active document) of the **com.sun.star.frame.Desktop** service without needing to explicitly create an instance of this service using the Service Manager and then use the **getCurrentComponent** method. 

Depending on where the macro is stored, **ThisComponent** represents either the containing document or the active (foreground) document (see Remarks). The properties and methods available through **ThisComponent** depend on the document type.

Syntax
------

``ThisComponent``

Examples
--------

.. code-block:: Basic
   
   ' Declarations
   
      Dim oDoc As Object

   ' Get a reference to a the containing or active document into oDoc
   
      oDoc = ThisComponent
   

Remarks
-------

If ThisComponent is used in a macro that is stored in a specific LibreOffice document container, then ThisComponent always returns a reference to the containing document. 

However, if ThisComponent is used in a macro that is stored in the **System and User Basic Macro Storage ("My Macros & Dialogs")** container, then it always returns a reference to the currently active (foreground) document. 

This differs from UNO API method ``getCurrentComponent()``, which always returns a reference to the currently active (foreground) document, regardless of where the script is located or how it is executed. This method is straightforward and does not depend on the storage location of the macro or script.

.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7
