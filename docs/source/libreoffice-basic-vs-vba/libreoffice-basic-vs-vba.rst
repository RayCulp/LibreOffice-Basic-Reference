A quick comparison of LibreOffice Basic and VBA
===============================================

If you have some experience creating macros with `Microsoft Visual Basic for Applications (VBA) <https://learn.microsoft.com/en-us/office/vba/api/overview/>`_ you will notice many similarities when you take your first look at LibreOffice Basic. However, this initial impression is deceptive. The LibreOffice UNO API is very different from the Microsoft Office object model, and the LibreOffice Basic IDE regrettably lacks many feature that the VBA IDE has. 

Here is a quick summary of what a VBA programmer can expect when taking their first steps with LibreOffice Basic.

Language syntax and features
----------------------------

LibreOffice Basic syntax and language fatures are almost identical to VBA, so in this regard VBA programmers should feel right at home.

Notable differences include LibreOffice Basic's lack of (native) support for user-defined enumerations and class modules. However, LibreOffice Basic includes two compiler options that add different levels of support for features that are otherwise specific to VBA: 

- ``Option Compatible``: Extends LibreOffice Basic compiler and runtime, allowing supplemental language constructs to Basic.
- ``Option VBASupport 1``: Allows LibreOffice Basic to support some VBA statements, functions and objects.

The LibreOffice IDE
-------------------

The LibreOffice Basic IDE is rudimentary and lacks support for many important features compared to the VBA IDE. 

When the rich feature sets provided by VBA IDE extensions such as `RubberDuck <https://rubberduckvba.com/>`_ and `MZTools <https://www.mztools.com/>`_ are thrown into the equation, writing code in the LibreOffice Basic IDE may seem hardly better than writing it in a simple text editor. 

The LibreOffice UNI API
-----------------------

The LibreOffice UNO API is very different from the Microsoft Office object model. VBA programmers familiar with the Microsoft Office object models should not expect to be able to leverage hardly any of their existing knowledge.

.. toctree::
    :titlesonly:
    :glob: