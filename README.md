# RUandSS_collector
Short description:
A simple program that, for a given product, find registration documents for it and create a document with the manufacturer and country of production.


Full description:
The repository is the program code itself (program_code), which is subsequently compiled into one whole file; template for choosing a base (Base template.xlsx); See this readme (README.md). The program is a pet project for solving a local problem of a Russian company (that is why Russian words are found in the code).

What the program can do:
Using the program, you can create a word file with a table that lists the specified positions of goods, their manufacturers and their country of production. In addition, based on the database data, the program can collect files associated with a given position in the database into an archive.

Instruction:
Work with the program is carried out step by step:
1. In a separate Excel file, a column contains the positions, information about which you need to find out. The path to this file is indicated via the "Specify Positions (.xls)" button.
2. The path to the pre-filled Base is indicated. Associated files in the Base must contain the name and extension (example.pdf). The files themselves must be located in the /RUandSS/example.pdf program folder. Several files are listed separated by commas and located in one cell of the table (example1.pdf, example2.pdf).
3. The Output Directory is indicated.
4. After stage 3, the green buttons light up, allowing you to create a Word document with the country of production and manufacturer, as well as an archive with associated documents.

