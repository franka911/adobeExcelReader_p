Script that reads and process data from 2 files: Excel and CVS ( x rows and 2 columns)
The script should do:
1. Scan whole CVS file
2. After encountering the first text frame in the CVS file, it should start looking for the same text in Excel in the first column,
 then when it finds it, it copies the contents of the same row, but from the second column and pastes it into the CVS
 text frame (such a text substitution)
3. After finishing the work, the program highlights in green those lines that it managed to find and replace in the project, 
and in red those that it did not find.
4. Program ignores white signs
