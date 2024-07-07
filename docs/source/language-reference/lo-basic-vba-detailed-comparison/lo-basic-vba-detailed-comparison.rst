Detailed comparison of LibreOffice Basic and Visual Basic for Applications
==========================================================================

The language
------------

.. list-table::
   :widths: 25 25 25
   :header-rows: 1
   :class: tight-table

   * - 
     - VBA
     - LibreOffice Basic
   * - **Syntax**
     - Similar to Visual Basic 6.0.
     - Similar to VBA but with some differences.
   * - **Data types**
     - Supports basic data types (Integer, Long, Double, String, etc.), complex types (Objects, Arrays, Collections) and classes.
     - Supports similar basic data types as VBA with some notabe exceptions, such as lack of (native) support for user-defined enumerations and class modules (support for some features specific to VBA can be enabled using the compiler options ``Option Compatible`` and ``Option VBASupport 1``)
   * - **Error handling**
     - Uses On Error GoTo and Err object, which has various methods and properties.
     - Uses On Error GoTo and Err function, which returns the error number as an Integer.
   * - **User forms**
     - Supports creating user forms with various controls (buttons, text boxes, etc.).
     - Supports dialogs and controls similar to VBA user forms, but the methods to instantiate and display a dialog. are different
   * - **Extensibility**
     - Highly extensible using COM libraries provided by other programs (e.g., Corel Draw, Adobe Acrobat), custom DLLs, OCX dialog controls, Declare Function statements, REST APIs using the full range of HTTP requests (GET, POST, PUT, DELETE).
     - Mostly limited to the UNO (Universal Network Objects) API for accessing LibreOffice components. Provides limited extensibility via custom UNO services, REST APIs using the HTTP Get request, and external command line executables.

The IDE
-------




.. toctree::
    :titlesonly:
    :glob: