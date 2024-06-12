# Flat XML ODF Spreadsheets Minimizer

Python 3 script that cuts out all unnecessary and optimizes styles in FODS documents
to make it much easier to solve conflicts of merging in version control systems.

Tested with _LibreOffice Calc 24.2.0_.

### Usage

Launch it in the directory that contains `.fods` documents.

### Features and limitations

- Deduplication, sorting, and stable names of styles (based on their hash).
- Cuts out all styles except the color of cells text and the color of cells background.
- Preserves frozen columns.
- Reduces documents size by ~60%.
- At the moment, it doesn't accept any arguments, and minimizes all the `.fods` documents that will be found in its working directory.
- Minimization rules have been found experimentally, so, make sure that after the minimization, your documents don't lose any useful information.
- Thanks to the FODS format, you can also use global formulas for entire columns. For example, in _Libreoffice Calc_, they can be added through menu `Format > Conditional > Manage... > Add > Formula is`:
  - To check all the values ​​of the column `A1` for uniqueness:  
    `COUNTIF($A$1:$A$1000, A1) > 1`
  - To check that all the values ​​of the column `A1` contain only allowed characters:  
    `LEN(REGEX(A1, "[0-9a-zA-Z]+", "")) > 0`

  So, such formulas will be also preserved.

### Motivation

Spreadsheets (electronic tables) are well suited for, for example, organizing a dictionary of localization of the application. The localization dictionary, inherently, is a text tabular document that is useful to store in versions control systems. But, depending on the chosen format of the document, there may be no certain possibilities, or problems arise when solving conflicts of merging.  
Let's look at the available Spreadsheets formats, without counting of formats with a special macros support:

| format | type     | formulas | stylization | additional
| :----- | :------- | :------- | :---------- | :----
| FODS   | **text** | **yes**  | **yes**     | Contains a lot of excess meta-information
| HTML   | **text** | no       | **yes**     | If the document contains several tables, _LibreOffice Calc_ doesn't want to open it
| CSV    | **text** | no       | no          | The easiest parsing format
| SLK    | **text** | no       | no          |
| ODS    | binary   | **yes**  | **yes**     |
| XLSX   | binary   | **yes**  | **yes**     |
| XLS    | binary   | **yes**  | **yes**     | The most difficult parsing format

Therefore, the only universal format suitable for the posed requirements is FODS. It allows you to store the text, and the stylization of the text, and the formulas, in a text format based on XML.

But, unfortunately, its implementation in _Libreoffice Calc_ contains a number of flaws:
- Deduplication of styles is not performed, that is, for different cells with one color, individual styles can exist. It increases the size of the document.
- The maintenance of stable names of styles is not carried out, as a result of which, the renaming of styles from the simple resaving of the document may occur. It increases the number of changes in the document.
- Automatic removal of unnecessary meta-information, which has default values, is not performed. It increases the size of the document.

Conclusion: In order to get a minimum size document, with looking at the specifics of use of Spreadsheets for, for example, localization dictionaries, it is much easier to write a similar script, rather than wasting time to improving of _LibreOffice Calc_ in the form of adding lots of options.
