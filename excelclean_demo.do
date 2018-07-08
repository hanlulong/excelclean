/*******************************************************************************


Demo of excelclean 

Author: Lu Han 
Last update: 08 Jul 2018

*******************************************************************************/

global home "https://github.com/hanlulong/excelclean/raw/master" 


di "Example 1: Oragnising downloaded excel files from Bankscope"

capture mkdir "demo_bankscope"
di "Copying Example Files"
copy "${home}/examples/Bankscope/A.xlsx" "demo_bankscope/A.xlsx"
copy "${home}/examples/Bankscope/B.xlsx" "demo_bankscope/B.xlsx"
copy "${home}/examples/Bankscope/C.xlsx" "demo_bankscope/C.xlsx"

excelclean , datadir("demo_bankscope") sheet("Results") cellrange("B1")  ///
             pivot integrate wordfilter(`"" Quarter""')  droplist("NID TONID")

di "All Done! The integrated dataset is saved under demo_bankscope/clean.dta"			 

			 
di "Example 2: Oragnising downloaded excel files from SNL"

capture mkdir "demo_SNL"
di "Copying Example Files"
copy "${home}/examples/SNL/A.xlsx" "demo_SNL/A.xlsx"

excelclean , datadir("SNL") sheet("Results") cellrange("A1") pivot namerange("1 3")
              
di "All Done! The integrated dataset is saved under demo_SNL/clean.dta"			 
