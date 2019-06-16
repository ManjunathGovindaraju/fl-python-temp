## Input Configuration


> input.yaml [Mandatory]:

File to specify the page number, column number, row number and file name . This will be input configuration file for the program

    number: 1                   - Excel Sheet number
    col : 'c'                   - Column index in small letters ( a to z)
    row : 5                     - Row index
    filename : "file01.txt"     - target file name. The content of the above mentioned section will be saved in this file
                                  (always overwritten)


## Usage

    python program.py <excel_file_name>
    
## Example  

    python program.py test001.xlsx



