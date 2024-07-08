CreateUnoService function
=========================

Creates an instance of a UNO service. The Basic function **CreateUnoService** is a shorthand way of creating an instance of a service without needing to use the Service Manager.

Syntax
------

``CreateUnoSerice`` ( ``_ServiceName_`` )

Parameters
----------

.. list-table::
   :widths: 25 25 25 25
   :header-rows: 1
   :class: tight-table

   * - Name
     - Required/Optional
     - Data type
     - Description
   * - *_ServiceName_*
     - Required
     - String
     - The name of the UNO service to be created

Return value
------------

Object. Returns a reference to the instance of the UNO service that has been created.

Examples
--------

.. code-block:: Basic
   
   Sub ListOpenDocuments
    
      ' Declarations

         Dim oDesktop As Object
         Dim oComponents As Object
         Dim oComponentEnum As Object
         Dim oDocument As Object
         Dim sTitle As String
    
      ' Create the Desktop service

         oDesktop = CreateUnoService("com.sun.star.frame.Desktop")

      ' Get the array of all open components (documents)

         oComponents = oDesktop.getComponents()

      ' Create an enumeration to iterate over the components
   
         oComponentEnum = oComponents.createEnumeration()

      ' Iterate over each component
   
         Do While oComponentEnum.hasMoreElements()

         ' Get the next component (document)
       
            oDocument = oComponentEnum.nextElement()

         ' Check if the document has a title property

            If HasUnoInterfaces(oDocument, "com.sun.star.frame.XModel") Then

           ' Print the document's title
            
               sTitle = oDocument.Title

               MsgBox "Document Title: " & sTitle

            End If

         Loop

   End Sub




.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7

.. 
   Todo:
