# polinate
A crudely made tool for cross-populating column data in multiple Excel files (since vlookup was taking too much time).

# demo
Working demo available [here](https://d-g-ivanov.github.io/pollinate/).

Test files available under example folder.

# details
The intended purpose of the tool is to take a given column (source) from multiple excel sheets, and cross-reference their text value to extract information from another column (target), same row. 

If there is no exact matchiing (including case and spacing) within source cells, the tool can use string similarity check to find fuzzy matches in the source to extract possible target.

Essenitally, it is vlookup within multiple files.

**NOTE:** Fuzzy target cell will have a comment added to them. The comment will contian:
- a string diff for the two sources showing the additions and deletions that happened
- a perentage of similarity between the 2 source strings
- the source string whose target was used

**NOTE:** Once matching is done, the browser will spit out the updated files. You might get a message to allow multiple downloads, please allow them. The new files will have the word “merged_” at the front for the purposes of distinguing them. Eventually, I might add a zip download option.

# configuration
All should be self-explanatory, but here is a brief manual.

**Worksheet name** – exact name of the excel tab where the data can be found
**Source cell** – the first cell in a column where the source text can be found
**Target cell** – the first cell in a column where the target text can be found
**Exact match in** – once a target is inserted, the cell will be colored to show that this text was added by the tool. Pick a color, or use the default one - greenish.
**No match in** – if a given source text is not found in any of the files that are analyzed, the target cell will be colored in this color. This way you know what is missing. Pick  a color, or use the default one – reddish.
**Use fuzzy** – whether or not the fuzzy algoritm should be used.
**Fuzzy match in** – if a fuzzy source match is found, the target cell will be colored in this color. This way you know the value is only a fuzzy match. Pick  a color, or use the default one – blueish.
**Min. fuzzy rating** – the acceptable string similarity rating when using fuzzy matching. It should be a value between 0 and 1.

# possible enhancements
- cannot match data from multiple sheets and columns at the same time. Works only for 1 sheet, with 1 source and 1 target column at a time. You can use it multiple times on already matched files though, just change the settings. Could possibly extend the functionality in the future.

# projects used
Will follow soon.

# licence
See the files.
