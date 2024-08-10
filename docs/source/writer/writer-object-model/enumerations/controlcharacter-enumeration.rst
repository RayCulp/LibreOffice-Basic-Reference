com.sun.star.text.ControlCharacter Enumeration (Writer)
========================================================

Defines the type of control character inserted by the :doc:`/writer/writer-object-model/swxbodytext-object/methods/insertcontrolcharacter`.

.. list-table::
   :widths: 25 25 25
   :header-rows: 1
   :class: tight-table

   * - Name
     - Value
     - Description
   * - PARAGRAPH_BREAK
     - 0
     - Starts a new paragraph.
   * - LINE_BREAK
     - 1
     - Starts a new line in a paragraph.
   * - HARD_HYPHEN
     - 2
     - Equivalent to a dash, but prevents this position from being hyphenated. TODO: Verify this.
   * - SOFT_HYPHEN
     - 3
     - Defines a special position as a hyphenation point. If a word containing a soft hyphen needs to be hyphenated at the end of a line, then this position is preferred.
   * - HARD_SPACE
     - 4
     - Links two words and prevents this concatenation from being hyphenated. It is printed as a space.
   * - APPEND_PARAGRAPH
     - 5
     - Aopends a new paragraph. TODO: Verify.

.. toctree::
   :titlesonly:
   :glob:
   :maxdepth: 7