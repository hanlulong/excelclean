# excelclean
A Stata module that extends the default `import excel` command and adds support on batch processing of excel files. 

## Features

* load all excel files in the specified directory
* organize variable names and labels
* reshape the dataset if necessary
* integrate all files into a cleaned dataset. 

The major improvement is on handling variable names and automating label generations. The module can manage most name exceptions in batch processing, including long variable names, multiline names, line breaks, special characters, etc. It generates variable names that can easily work with the `reshape` command and provides `pivot` and `integrate` options to automate the entire process. 

## Usage 
0. Installation 

          net install excelclean, from("https://github.com/hanlulong/excelclean/raw/master")

1. Put all excel files of similar kinds into the same directory 

2. Run 
          
          excelclean, datadir(string) sheet(string) cellrange(string) [options]

### Example Usage 
After installing the package, run

          excelclean_demo


## Syntax

    excelclean, datadir(string) sheet(string) cellrange(string) [options]

        datadir(string) 
            directory where excel files are stored. Please close and save all excel files 
            under this directory before executing the command.
            e.g., datadir("c:/myplace/")

        sheet(string) 
            the excel sheet in each excel file to be loaded into Stata.
            e.g., sheet("sheet1") or sheet("Results")

        cellrange(string) 
            the range of data cells on each excel sheet to be loaded into Stata.
            e.g., cellrange("A1") to extract all cells; cellrange("B3") to extract cells
            starting from the second column and the third row.


## Options 

        integrate 
            integrate all datasets into a single dta file. The default is to save each dta file
            separately using the name of the corresponding excel file.

        pivot 
            reshape variables from a wide format to a long format. By default, the program
            recogonizes the last word from the formulated variable names as the time indicator.
            The program will detect the variables that need to be reshaped into a long format, e.g.
            Var2000, Var2001, Var2002 -> Var

        droplist(string) 
            drop redundant variables from the dataset in the data integration process. It
            helps to reduce the file size and processing time. Separate variable names by a space " ".
            e.g., droplist("Var1 Var2 "). Remember to leave a space at the end of the last variable.

        resultdir(string) 
            specify the directory where the results are saved. The default is the directory "datadir" 
            where the excel files are stored.
            e.g., resultdir("C:/myresultdir/"). 
            Always use the full path of the directory to aviod possible conflicts.

        extension(string) 
            specify the extension of files to be included. The default is "xlsx".
            e.g., extension("xls")

        namerange(integer) 
            specify the rows that record variables names. The default is the first row.
            e.g., namelines(4) to indicate that the first four rows contain information about variable
            names. Note that the first four lines will be deleted from the dataset after creating the
            variable names.

        namelines(string) 
            select the rows to formulate variables names. The default is the first row.
            e.g., namelines("1 3") to specify the first and the third rows to be used to formulate
            variable names

        wordfilter(string) 
            specify specific characters to be excluded from variable names.
            e.g., wordfilter("year quarter the"); to exclude the space before any word use
            wordfilter(`"" word1" " word2""')

