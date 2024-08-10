SwXBodyText.insertControlCharacter method (Writer)
====================================================

Inserts a control character into the text. Control characters in the LibreOffice Writer object model are non-printable characters that control the formatting and layout of text within a document. These characters influence how text is displayed and how various document elements behave, but they are typically invisible during normal text editing.

Syntax
------

*expression*. ``insertControlCharacter`` ( ``_TextCursor_`` , ``_ControlCharacter_`` , ``_Absorb_``)

*expression* A variable that represents an :doc:`/writer/writer-object-model/swxbodytext-object/swxbodytext-object`.

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
   * - *_TextCursor_*
     - Required
     - Object of type Range or Text Cursor
     - A text cursor or a text range object that specifies the position in the text where the control character will be inserted.
   * - *_ControlCharacter_*
     - Required
     - Short
     - Must be of the predefined constants from the com.sun.star.text.ControlCharacter enumeration.
   * - *_Absorb_*
     - Required
     - Boolean
     - Determines whether the inserted control character replaces the current text selection (if any). If ``_Absorb_`` is True, any text currently selected by ``TextCursor`` will be replaced by the control character. If bAbsorb is False, the control character is inserted without deleting any text, and the current text selection remains unchanged.

Return value
------------

None

Examples
--------

.. code-block:: Basic

   Sub InsertParagraphBreak
   
      ' Declarations

         Dim oDoc As Object
         Dim oText As Object

      ' Get the document and text

         oDoc = ThisComponent
         oText = oDoc.Text

      ' Create a text cursor at the beginning of the document
         
         oTextCursor = oText.createTextCursor
      
      ' Insert a paragraph break at the cursor position

         oText.insertControlCharacter(oTextCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
   
   End Sub


.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7