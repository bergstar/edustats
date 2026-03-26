# Prompt History

Representative prompts used to build this repository, annotated with the matching commits.

## Prompt 01

**Commit:** `14962f25`  
**Message:** `Add Excel processing pipeline for Russian educational reports`

> now we need our first Python script to process actual files.
>
> for this we need to first refactor our file structure.  
> files are named in this scheme  
> г.Москва_ГОС_Очная  
> (region)_(governmental)_(format)  
> which represent education region, governmental or commercial, and format itself.
>
> using regexp refactor and restructure files so that we have following structure  
> format/governmental/region
>
> translate Russian words into English for the directory structure
>
> keep the region names as is, convert all to lower case
>
> call this script 001.py

## Prompt 02

**Commit:** `14962f25`  
**Message:** `Add Excel processing pipeline for Russian educational reports`

> great, now Lets move on to create our second script and second output.
>
> for this we need to export each sheet from each of the Excel documents and store them as separate files.
>
> for this we dont need to re-copy or store the original .xlsx file in the new output folder
>
> keep only the sheet names as-is in the document. for this we will assume they are uniq
>
> since there is an issue with "copy" of the sheets to a new file, we will go the other way around, and delete all  
> the sheets we dont need, keep only the sheet we need and store it as a single-sheet file. for this we will use  
> the openpyxl library

## Prompt 03

> make the script only run for the first book. we dont need it to run for all

## Prompt 04

> and debug output to the script, remove limit of only 1 book, also remove code which is related to checking if  
> anything exists. it should delete and recreate folder on each run.  
> dont run the script  
> remove all parameters - like dry run etc.

## Prompt 05

**Commit:** `233d7bf8`  
**Message:** `Add parallel processing to Excel sheet export`

> output is extremely messy, clean it up, no need to spam that much. only output what is needed to see real  
> progress. also use some ASCII output / terminal / progress library to make it nice. no emojis.
>
> its very slow now, we need to run it in parallel so it parses at least 10x faster
>
> introduce an extremely minimalistic parallel processing - KISS - dont go crazy

## Prompt 06

**Commit:** `02d8f50f`  
**Message:** `Add real-time progress tracking to Excel sheet export`

> its okay but terminal updates too slow. make it update faster so its a bit more fun to watch. maybe showcase  
> number of sheets - let it rip!!

## Prompt 07

> great make the number of workers a terminal input parameter. set default to 8. give me command to run with 12

## Prompt 08

**Commit:** `da489a3d`  
**Message:** `Add header extraction step to Excel processing pipeline`

> now we need to write our third script, this script will work on cleanup of the files.
>
> each Excel document has from 1 to 10 lines in the header of the documents. those lines are merged full-width  
> cells - spanning full width of the document. for each document, we need to re-copy it to the new  
> destination and delete those lines. again delete and preserve the formatting. store the values of those cells  
> in JSON with output lines 01, 02, 03, 04 etc. also export and store empty lines. everything string. KISS

## Prompt 09

> we got an error since one of the files has an additional empty column which messes up the algorithm.
>
> Lets check for totally empty columns from the right, and delete them if none of the cells are merged and there  
> is no data.
>
> ➜  study git:(master) ✗ python3 003.py  
> Cleaning 18438 files with 8 workers...  
> [--------------------------------] 33/18438 files | 190 header lines | 121.9 files/s | 00:00:00  
> Failed while cleaning /Users/dk/Development/study/002_output/full_time/commercial/алтайский край/Р3_3_1.xlsx:  
> Expected 1 to 10 header rows, got 0

## Prompt 10

**Commit:** `eff1c89f`  
**Message:** `Add workbook merge step to Excel processing pipeline`

> great now Lets create our 004 script.
>
> this time we need to merge sheets
>
> so files like Р2_1_1 Р2_1_1(2) Р2_1_1(3) etc are all parts of the same larger sheet which was split into  
> smaller chunks. we need to merge them back
>
> for this we will only output Р2_1_1 which will contain all the merged files of that range.
>
> all chunks of the larger table have the same starting columns, they are the same in all chunks. usually its at  
> least 1 but not more than 7.
>
> we will keep the description first columns as-is in the first file, but skip / avoid copying them from  
> subsequent files.
>
> so in other words in our files we have a number of columns from A to G, which are the descriptive columns in  
> all files. we are interested in copying and adding columns from other files to the right of existing data in  
> original file, while avoiding copying descriptive columns. data should be sequentially to the right of the  
> sheet.
>
> since there is a chance that there is some space or other minor typos, before comparing column names, clean  
> them up and keep only the а-яА-Я0-1 characters, so we minimize the risk of misidentification.
>
> since there is a chance of misidentification, also do a CRC32 sum of raw data dump of the first 5 rows down  
> from the header row. please note that header row might be a merged cell, so we of course only care about the  
> data from the 3 non-merged cells below. if less than 3 cells below the header exist, skip this step.

## Prompt 11

> we dont need to try merging files which are not split. so if no (2) exists we just copy that file without change
>
> ➜  study git:(master) ✗ python3 004.py  
> Merging 12731 output files...  
> [--------------------------------] 1/12731 files | 1 parts | 1056.7 files/s | 00:00:00  
> Failed while merging /Users/dk/Development/study/003_output/full_time/commercial/алтайский край/Р1_2.xlsx:  
> CRC32 mismatch in descriptive rows

## Prompt 12

> let it run even if there is an error, output errors to debug.txt so that you can study them later on.

## Prompt 13

**Commit:** `fb25684e`  
**Message:** `Skip empty rows and section delimiters in CRC hash calculation`

> work on solving the errors according to the debug data in debug.txt
>
> make a plan before going forward with implementation

## Prompt 14

> /Users/dk/Development/study/004_output/full_time/governmental/г.москва/Р2_1_2.xlsx
>
> here we got an error. first cell A1 spans 4 rows in Р2_1_2.xlsx  
> however in Р2_1_2(4).xlsx it spans 5 rows
>
> we need to check if header cells in all of documents span more cells than what is in the main file. if that  
> is the case, we need to add a blank row to the main document if it has fewer rows than what is in subsequent  
> files, and if it has more - then add blank row to the files with fewer rows in the merged header cells.
>
> algorithm here is to identify the most spanned header cell in height, take it as a header height for that chunk,  
> and adjust accordingly

## Prompt 15

**Commit:** `ebbf6ed4`  
**Message:** `Normalize header heights during workbook merge`

> algorithm is wrong. it added an additional row to the first document at the bottom, then added a row on top of the  
> second document, again at the bottom etc.
>
> review your logic.
>
> again use /Users/dk/Development/study/004_output/full_time/commercial/г.москва/Р2_1_3.xlsx as example where  
> this is the case

## Prompt 16

> now we need to move forward to the actual conversion of data to the JSON format.
>
> now we will be working on creating script 005
>
> all headers consist of advanced structured multi-merged and spanning cells which contain description, to one or  
> several rows - grouped at different levels. but always following down to a specific single row at the lowest  
> end. after that there is a level of columns with running numbers running from 1 and until the end.
>
> we need to convert that upper header structure into a structured lookup file, where we go bottom up
>
> /Users/dk/Desktop/Screenshot 2026-03-24 at 16.07.31.png
>
> again use /Users/dk/Development/study/004_output/full_time/commercial/г.москва/Р2_1_3.xlsx as example where  
> this is the case
>
> So for row 15 we will have this structure
>
> - 5: на места в пределах отдельной квоты приема
> - 4: из суммы гр. 10, 12, 13 - поступившие
> - 3: за счет бюджетных ассигнований
> - 2: обучались
> - 1: В том числе (из гр.5)
>
> and for row 8 we will get something like
>
> - 3: специали-тета
> - 2: продолжили обучение по программам
> - 1: В том числе (из гр.5)
>
> For cases where we have vertically merged cells, we will use the lowest number of the merged cell. so if 3+4+5  
> is merged - we refer to it as 3.
>
> we save it as Р2_1_3_library.json etc
>
> we remove headers with text for the output xlsx files
>
> we do this even if files are empty or blank
>
> we keep the row with running column numbers

## Prompt 17

**Commit:** `4ebb70d2`  
**Message:** `Add library extraction and header removal step to pipeline`

> I dont like this
>
> "Deleting rows in place keeps the cell formatting, but the old merged header ranges stay behind and corrupt the  
> top of the sheet. I’m testing a lighter fix now: strip/rebuild merges around the delete instead of copying  
> every cell by hand."
>
> I want you to find the rows which are the header rows with text above the running number row we  
> need to keep. after finding the edges, you need to remove those rows entirely all at once.

## Prompt 18

**Commit:** `63eb07d6`  
**Message:** `Add SQL bundle generation step to Excel processing pipeline`

> Design the 005 -> 006 conversion so that each cleaned workbook becomes one wide SQL-ready dataset table plus  
> two dictionary tables. Use the naming model F/H/P + _ + C/G + _ + region code + _ + form code.  
> Region codes coming from a shared regions lookup.
>
> For full_time/commercial/г.москва/Р2_1_3.xlsx, your bundle would be:
>
> - Main table: F_C_08_Р2_1_3
> - Column dictionary: F_C_08_Р2_1_3_column
> - Row dictionary: F_C_08_Р2_1_3_row
> - Shared lookup: regions
>
> regions
>
> | region_code | region_name |
> | 08          | г.москва.   |
>
> each table should be its own .sql file all put into a single folder
>
> there should be a simple import script which asks for the table name and imports all the subsequent files into the db.  
> it will also ask for user name and password - from terminal - --user --password --database

## Prompt 19

**Commit:** `b393172d`  
**Message:** `Add database import step to pipeline`

> before importing drop all tables in that database  
> make parameters from the command line like  
> --database data --user data --password data  
> make localhost a default host which can be overridden with --host

## Prompt 20

**Commit:** `f9779db6`  
**Message:** `Add regional workbook merging step to pipeline`

> now that we have a solid 006 and 005 script working, Lets create a 007 script - which will take 004 as input  
> but now will merge all regions together. of course we must keep the full_time / commercial etc structure as is  
> but functionally it will be like 005 - but a bit more advanced, while keeping the basic structure  
> for the merge we will take the first regional file, add a new column (last) lets call it Region  
> here we will populate it with the region number  
> then all subsequent regions will get inserted under the first table. beware of the Справка block, it has to  
> be removed in this. also the header is only kept once, same with running column numbers - we only keep that  
> for the first Excel file.
>
> so basically we need to merge all the regions into single files, we need to do it clean, while preserving the  
> header of the first instance.
>
> it should output in same format as original 005 script, but this time with merged files. output to 007

## Prompt 21

**Commit:** `377ff6ee`  
**Message:** `Preserve zero-padded formatting when reading Excel cells`

> Lets now go back to the 006 script, it takes values from "№ строки" like 01 02 03 but stores as 1, 2, 3  
> change it so its identical to what is in the files  
> if 01 store as 01

## Prompt 22

**Commit:** `4b80ebdd`  
**Message:** `Add workbook-to-SQL conversion step to pipeline`

> lets now create the 008 script which will do the same as 006, but will combine what 006 and 007 does into a  
> single script. basically combined SQL conversion with all regions into single tables as-is with 007
