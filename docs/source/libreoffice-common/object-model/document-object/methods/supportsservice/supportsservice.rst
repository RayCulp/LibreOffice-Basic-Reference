Document.SupportsService method (LibreOffice common)
=========================================================

Determines whether the Document object is of the type specified by the ``_ServiceName_`` parameter, for example a Calc spreadsheet, a Writer document or an Impress presentation.

Syntax
------

*expression*. ``SupportsService`` ( ``_ServiceName_`` )

*expression* A variable that represents a Document object. Can also be the ``ThisComponent`` object.

Parameters
----------

Parameters
++++++++++

.. list-table::
   :widths: 25 25 25 25
   :header-rows: 1

   * - Name
     - Required/Optional
     - Data type
     - Description
   * - *_ServiceName_*
     - Required
     - String
     - The name of the Service that the Document object should be tested for.

Service names (document types)
++++++++++++++++++++++++++++++

.. list-table::
   :widths: 25 25
   :header-rows: 1

   * - Name
     - Document type
   * - com.sun.star.sdb.OfficeDatabaseDocument
     - Base
   * - com.sun.star.sheet.SpreadsheetDocument
     - Calc
   * - com.sun.star.drawing.DrawingDocument
     - Draw
   * - com.sun.star.presentation.PresentationDocument
     - Impress
   * - com.sun.star.text.TextDocument
     - Writer

Return value
------------

Boolean. Returns ``True`` if the Document object is of the type specified by the _ServiceName_ parameter, returns ``False`` if it is not.

Remarks
-------

To avoid runtime errors, this method should be used to determine the type of the Document object before executing other methods or querying properties that are specific to an individual document type.

Examples
--------

.. code-block:: Basic
   
   ' Declarations
   
      Dim sDocumentType As String 
      
   ' Use the SupportsService method to check if the document supports a specific service name
   ' The service name indicates the type of the document, such as text, spreadsheet, presentation, or database

      Select Case True
      
      Case oDoc.SupportsService("com.sun.star.sdb.OfficeDatabaseDocument")
      
         sDocumentType = "BASE"
         
      Case oDoc.SupportsService("com.sun.star.sheet.SpreadsheetDocument")
      
         sDocumentType = "CALC"
         
      Case oDoc.SupportsService("com.sun.star.drawing.DrawingDocument")
      
         sDocumentType = "DRAW"
         
      Case oDoc.SupportsService("com.sun.star.presentation.PresentationDocument")
      
         sDocumentType = "IMPRESS"
         
      Case oDoc.SupportsService("com.sun.star.text.TextDocument")
      
         sDocumentType = "WRITER"
         
      Case Else
      
         GetDocumentType = "UNKNOWN"
         
      End Select


.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7