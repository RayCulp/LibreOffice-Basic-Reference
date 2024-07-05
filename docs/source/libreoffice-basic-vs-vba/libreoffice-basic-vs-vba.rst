LibreOffice Basic vs. VBA
=========================

Programmers familiar with `Microsoft Visual Basic for Applications (VBA) <https://learn.microsoft.com/en-us/office/vba/api/overview/>`_ will immediately notice many similarities when they begin to program in LibreOffice Basic. However, there are also some key differences. Here is what a VBA programmer should expect when learning to program LibreOffice Basic.

The language
------------

LibreOffice Basic language features and syntax are almost identical to VBA, so VBA programmers should feel right at home. 

Some notable differences include LibreOffice Basic's lack of (native) support for user-defined enumerations and class modules. 

However, LibreOffice Basic includes two compiler options that add two different levels of support for features that are otherwise specific to VBA: 

- ``Option Compatible``: Extends LibreOffice Basic compiler and runtime, allowing supplemental language constructs to Basic.
- ``Option VBASupport 1``: Allows LibreOffice Basic to support some VBA statements, functions and objects.

The IDE
-------

The LibreOffice Basic IDE is rudimentary and lacks support for many important features compared to the VBA IDE. 

When the rich feature sets provided by VBA IDE extensions such as `RubberDuck <https://rubberduckvba.com/>`_ and `MZTools <https://www.mztools.com/>`_ are thrown into the equation, writing code in the LibreOffice Basic IDE may seem hardly better than writing it in a simple text editor. 

The LibreOffice object model
----------------------------

The LibreOffice object model is completely different from the Microsoft Office object model. VBA programmers familiar with the Microsoft Office object models should not expect to be able to leverage any of their existing knowledge.

.. toctree::
    :titlesonly:
    :glob: