# VBA macros for Excel
===

A little bit of code I wrote to improve my workflow with some functions I thought Excel was missing

### CopyPaste

The macros in this file use the computer's clipboard to store cell values instead of relying in Excel's clipboard.

The **CopyUniqueValues** macro only stores each value once, removing the duplicates.

The **PasteInRegions** macro checks if the number of cells selected is equal to the number of cells copied and pastes them in the same order.
This macro is specially helpful to paste values only in visible cells when a filter is applied.


### Filters

The macros in this file add some features to Excel's filtering options.

The **FilterByCopiedRange** macro filters by the values stored in the clipboard.

The **FilterByNotInCopiedRange** macro filters by the values not stored in the clipboard.

THe **CellsWithFilter** macro goes sequentially to the cells whith applied filters.
If mapped to a shortcut key, it is specially helpfull when working with files with a lot of columns.


### SelectionInfo

The macros in this file display some info about the selected cells in a Message Box.

The **UniqueValues** macro displays a message saying how many different values are in the selected cells and how many of them occur only once.

The **Product** macro displays the product of all the selected cells with numerical values.

The **Difference_and_Ratios** macro displays the difference and ratio between two selected cells with numerical values.


### ShortMacros

Short macros with some simple functions not included in Excel.

The **ChangeCase** macro changes between lowercase, UPPERCASE and Titlecase the text in the selected cells.

The **TrimText** macro removes all space characters at the beginning or end of the text in the selected cells.

