StarDesktop object
==================

The ``StarDesktop`` object is a predefined global variable that provides access to the ``com.sun.star.frame.Desktop`` service. It serves as a convenient shortcut to access the Desktop service without needing to explicitly create an instance of it through the Service Manager.

Syntax
------

``StarDesktop``

Remarks
-------

Both ``StarDesktop`` and the ``com.sun.star.frame.Desktop`` service accessed via the Service Manager ultimately refer to the same Desktop service. Using ``StarDesktop`` is a convenient shortcut for most macro tasks, while the Service Manager approach offers more explicit control and flexibility for advanced scenarios where you need to create other services that are not globally available as shortcuts.

Methods
-------

The methods of the ``StarDesktop`` object are identical to those of the com.sun.star.frame.Desktop service.

Properties
----------

The properties of the ``StarDesktop`` object are identical to those of the com.sun.star.frame.Desktop service.

Examples
--------

.. code-block:: Basic
   
   Sub ListAllOpenDocuments_StarDesktop()
      
      ' Declarations

         Dim oComponents As Object
         Dim oEnum As Object
         Dim oDoc As Object
         Dim sDocInfo As String

      ' Use the global StarDesktop variable to get the collection of all open documents

         oComponents = StarDesktop.getComponents()
         oEnum = oComponents.createEnumeration()
      
      ' Iterate through the collection and gather information

         sDocInfo = "Currently open documents:" & Chr(13)
         
         Do While oEnum.hasMoreElements()
         
            oDoc = oEnum.nextElement()

            ' Check if the object supports the XModel interface

            If HasUnoInterfaces(oDoc, ""com.sun.star.frame.XModel"") Then
                  sDocInfo = sDocInfo & oDoc.Title & Chr(13)
            End If

         Loop
      
      ' Display the information
      
         MsgBox sDocInfo

   End Sub
   

.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7
