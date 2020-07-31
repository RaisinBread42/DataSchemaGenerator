# DataSchemaGenerator

Used to generate C# classes and SQL Tables by tables defined in Excel worksheets.

## Motivation
During requirement analysis of data, be it table structure, checking quality of data, or entity relationships, Excel is often used as data can be shared in a presentable format.  After analysis, developers typically need to create relevant C# classes and SQL tables - mirroring the data within Excel. What if we had a tool to do this for us? It would save us time that can be used elsewhere.

## Capabilities
1. Generate C# classes with Id as key and other foreign keys detected.
2. Generate SQL tables with Id as IDENTITY (1,1), and adding foreign keys detected

Note - Currently, they're very little validation checks in place.
## RoadMap
*  Add seeding statements for C#
*  Add seeding for SQL tables

## How to Use

### Preparing Data
Simply create tables starting from cell A1, and remember the folloing rules. 
*  Ensure to set the correct data type for each column. General = nvarchar(200)/string, Number = INT/int. 
*  Don't need to add Id IDENTITY columns as the program will automatically add them in as primary keys. 
*  Worksheet names will be used for class and table names. Using nonalphabetical characters may result in the program not working correctly.
*  To add foreign keys, simply prepend FK_ followed by the name of the table/worksheet holding the references.

Please see DataSchemeGenerator/Templates/TestClassGenerationDataFile.xlsx as an example.

### Generating Schemas
Pretty straight forward. Upload the Excel file to be used, check whether you want C# classes generated, if so optionally enter namespace to use, and or SQL tables to be generated.

Upon clicking 'Show me the magic!' a Result.cs/Result.sql will open using the default associated program on your PC. If this fails to happen, please check the temp directory for the resulting file.
