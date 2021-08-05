# salty


Compute inter-rater reliability by comparing SALT files.

## Dependencies
Salty depends on the packages `chardet` and `xlsxwriter`.


## Usage

1. Run salty

   `python salty.py`

2. A file selection dialog will appear.  Select **two** .slt files to be compared.
3. A new window will appear showing the contents of the two selected files side by side.
The program will attempt to align the contents of the files, but you can manually adjust
alignment in this window.  See Manual Alignment Controls.  When alignment is complete, click the Done button.
4. A new window will appear showing the computed comparison between the SALT transcripts.  Values can be
manually edited by clicking in a cell to cyle between the values (1, 0, blank).  Right click to reset the
cell to the original computed value.  When editing is complete, click the Done button to generate the output
file containing the comparison.



## Manual Alignment Controls

### Row operations
- Ins: Insert a blank row above the selected row
- Del: Delete selected row(s)
- m: Mark selected row(s)

### Cell operations
- C-x: Cut selected cell(s), placing the contents into a buffer
- C-v: Paste the contents of the buffer, if any, into a range starting with the first selected row.

- C-z: Undo the last operation (cut/paste/delete/insert)  

Cell operations are performed *within a single column*.  

Click on any row (below the first row) to select it.  To select multiple rows,
hold the Control key while clicking on individual rows.  To select a range of rows,
select the first row in the range, hold the Shift key while clicking on the row at the
end of the range.


