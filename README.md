# Excel-Macros
this is an excel macro which I wrote to update the value of a cell.
It will update your cell value as
<code>
newValue = '+oldValue+',
</code>
Althugh this could be simply done with a simple command in excel also.
Suppose you have an excel column <code> A </code> which you need to update.
Select any colmun suppose you selected <code> B </code>
So you can right <code>B1</code> as 
<code>
B1 = "'"&A1&"',"
</code>
this will insert values of column<code>A</code> into column <code> B </code> with charater quote and coma separator.

Still you want to use this code then please see "how-to-use" file

