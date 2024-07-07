HasUnoInterfaces function
=========================

Tests whether a Basic UNO object supports one or more UNO interfaces passed as arguments.

Syntax
------

``HasUnoInterfaces`` ( ``Object``, ``Uno-Interface-Name-1`` [, ``Uno-Interface-Name-2`` , ...])

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
   * - *Object*
     - Required
     - Object
     - The Basic UNO object to be tested
   * - *Uno-Interface-Name-X*
     - Required
     - String
     - One or more UNO interface names (separated by commas) to test for

Return value
------------

Boolean. Returns ``True`` if **all** of the UNO interface names passed as arguments are supported, ``False`` if not.

Examples
--------

The following example tests whether the active documnt is a Writer document.

.. code-block:: Basic
   
   Sub CheckIfWriterDocument()

      ' Declarations

         Dim oDoc As Object
    
      ' Get a reference to the active document

         oDoc = ThisComponent
    
      ' Check if the document supports the XTextDocument interface
    
         If HasUnoInterfaces(oDoc, "com.sun.star.text.XTextDocument") Then
            MsgBox "This is a Writer document."
         Else
            MsgBox "This is not a Writer document."
         End If

   End Sub

The following example tests whether the active selection in a Calc document is a cell object.

.. code-block:: Basic
   
   Sub CheckIfCell()

      ' Declarations

         Dim oSelection As Object
    
      ' Get a reference to the current selection in the active document

         oSelection = ThisComponent.getCurrentSelection()
    
      ' Check if the selection supports the XCell interface
    
         If HasUnoInterfaces(oDoc, "com.sun.star.table.XCell") Then
            MsgBox "The selection is a cell."
         Else
            MsgBox "The selection is not a cell."
         End If

   End Sub



.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7

.. 
   Todo:
   Add list of possible interface names. This will probably be a massive list,
   so maybe it would be better to break it down by object type.