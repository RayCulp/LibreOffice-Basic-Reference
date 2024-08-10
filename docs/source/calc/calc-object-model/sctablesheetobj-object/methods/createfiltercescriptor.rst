ScTableSheetObj.CreateFilterDescriptor method (Calc)
====================================================

.. note::

   This article is a draft. The name of the parent object type "ScTableSheetObj" may not be correct.

Creates an object of the type ScFilterDescriptorBase that can be configured and passed as an argument to the Filter method of the ScTableSheetObj object.

Syntax
------

*expression*. ``CreateFilterDescriptor`` ( ``_Empty_`` )

*expression* A variable that represents an :doc:`/calc/calc-object-model/sctablesheetobj-object/sctablesheetobj-object`.

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
   * - *_Empty_*
     - Required
     - Boolean
     - If set to TRUE, then an empty filter descriptor is created. If set to FALSE, then the descriptor is filled with the previous settings of the current object (such as a database range).

Return value
------------

Object. An object of the type ScFilterDescriptorBase that can be configured and passed as an argument to the Filter method of the ScTableSheetObj object.

Examples
--------

.. code-block:: Basic
   
   ' Declarations

      Dim objSheet As Object
      Dim objFilterDescriptor As Object
      Dim objTableFilterField(0) As New com.sun.star.sheet.TableFilterField

   ' Get a reference to the sheet that you want to filter

      objSheet = ThisComponent.Sheets.GetByName("Sheet1")
      
   ' Create an empty filter descriptor object
   
      objFilterDescriptor = objSheet.CreateFilterDescriptor(True)
      
   ' Filter the sheet using the empty descriptor to remove any previously applied filters
   
      objSheet.Filter(objFilterDescriptor)
      
   ' Configure the filter descriptor

      With objFilterDescriptor                           
         .ContainsHeader = True         ' The sheet has headers         
         .UseRegularExpressions = True  ' Use regular expressions
      End With
      
   ' Configure the table filter field
   
      With objTableFilterField(0)
         .Field = 0                                          ' Filter on column A
         .IsNumeric = False                                  ' The value to filter for is a string, not a number
         .Operator = com.sun.star.sheet.FilterOperator.EQUAL ' Use the EQUAL operator
         .StringValue = ".*" & "text to look for" & ".*"     ' The expression to filter for
      End With   
      
   ' Set the FilterFields property of the Filter Descriptor to the Filter Field array

      objFilterDescriptor.setFilterFields(objTableFilterField())

   ' Apply the filter
      
      objSheet.Filter(objFilterDescriptor)   


.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7