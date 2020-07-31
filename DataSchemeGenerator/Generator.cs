using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Diagnostics;


namespace DataSchemeGenerator
{
    public static class GENERATION_METHODS
    {
        public const string CLASS = "CLASS";
        public const string SQL = "SQL";
    }

    public class Generator
    {
        public string FileName { get; set; }
        public string DocPath { get; set; }
        public string NamespaceToUse { get; set; }
        public string Contents { get; set; }
        public List<string> GenerationMethods { get; set; } = new List<string>();
        public string TempGenerationResultFileName { get; set; }

        public Generator()
        {
            // to avoid debug exception on new license for EPP 
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
        }

        public void GenerateSchema()
        {
            foreach (var method in GenerationMethods)
            {
                switch (method)
                {
                    case GENERATION_METHODS.SQL:
                        GenerateSQLSchema();
                        break;
                    default:
                        GenerateClassSchema();
                        break;
                }
            }
            
        }

        private void GenerateSQLSchema()
        {
            // get or define resources
            string resultFileName = Path.Combine(Path.GetTempPath(), "Results.sql");
            string newSQLTableTemplate = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Templates\\NewSQLTableTemplate.txt");
            string newSQLWrapperTemplate = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Templates\\NewSQLTableWrapperTemplate.txt");
            string tableResults = "";
            string SQLResults = "";

            //open excel data file, loop through each worksheet and read table columns
            using (ExcelPackage package = new ExcelPackage(new FileInfo(DocPath)))
            {
                int colCount = 0;
                int rowCount = 0;
                string worksheetTableWrapperGenerationResult = newSQLWrapperTemplate;

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    string worksheetTableGenerationResult = newSQLTableTemplate;
                    //get Column Count. Do not consider empty columns
                    //get row count. Do not consider empty rows
                    colCount = worksheet.Dimension.End.Column;  
                    rowCount = worksheet.Dimension.End.Row;

                    worksheetTableGenerationResult = worksheetTableGenerationResult.Replace("[TABLE_NAME]", worksheet.Name);
                    for (int row = 1; row <= 1; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var currentCell = worksheet.Cells[row, col];
                            string variableType = GetSQLColType(currentCell);

                            worksheetTableGenerationResult = UpdateSQLColumns(worksheet.Name, currentCell, worksheetTableGenerationResult);
                            worksheetTableGenerationResult = DefinePrimaryKey(worksheet.Name, currentCell, worksheetTableGenerationResult);
                            worksheetTableWrapperGenerationResult = UpdateSQLReferences(worksheet.Name, currentCell, worksheetTableWrapperGenerationResult);
                        }

                        //clean up placeholder
                        worksheetTableGenerationResult = worksheetTableGenerationResult.Replace("    [COL_NAME]\r\n", "");

                        tableResults = string.Concat(tableResults, worksheetTableGenerationResult);
                    }
                }

                //clean up
                worksheetTableWrapperGenerationResult = worksheetTableWrapperGenerationResult.Replace("[REFERENCES]", "");

                // create final SQL 
                SQLResults = worksheetTableWrapperGenerationResult.Replace("[TABLES]", tableResults).Replace("[REFERENCES]", "");

                //finally save file 
                File.WriteAllText(resultFileName, SQLResults);

                //open file for user using associated default application
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo(resultFileName)
                {
                    UseShellExecute = true
                };
                p.Start();
            }
        }

        #region SQL Generation Helpers

        private string UpdateSQLColumns(string worksheetName, ExcelRange cell, string worksheetTableGenerationResult)
        {
            string variableType = GetSQLColType(cell);

            if (cell.Text.Contains("FK_"))
            {
                // add column, append placeholder for next col to be added
                // add FK constraint 
                var varName = cell.Text.Replace("FK_", "");
                worksheetTableGenerationResult = worksheetTableGenerationResult.Replace("    [COL_NAME]", String.Concat("    [", varName, "Id] INT NOT NULL,", "\n    [COL_NAME]"));
                return worksheetTableGenerationResult;
            }
            else
                //replace placeholder with new prop, then put placeholder on new line for next prop
                return worksheetTableGenerationResult.Replace("    [COL_NAME]", String.Concat("    [", cell.Text, "] ", variableType, " NULL,", "\n    [COL_NAME]"));

            return string.Empty;
        }
        private string DefinePrimaryKey(string worksheetName, ExcelRange cell, string worksheetTableGenerationResult)
        {
            return worksheetTableGenerationResult.Replace("[PK]", String.Concat("CONSTRAINT [PK_",worksheetName,"] PRIMARY KEY CLUSTERED\n ( [Id] ASC) WITH ( PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] \n) ON[PRIMARY] \n GO"));
        }
        private string UpdateSQLReferences(string worksheetName, ExcelRange cell,string SQLWrapperGenerationResult)
        {
            string variableType = GetSQLColType(cell);

            if (cell.Text.Contains("FK_"))
            {
                // add column, append placeholder for next col to be added
                // add FK constraint 
                var varName = cell.Text.Replace("FK_", "");
                SQLWrapperGenerationResult = SQLWrapperGenerationResult
                    .Replace("[REFERENCES]", String.Concat("ALTER TABLE [dbo].[", worksheetName, "] WITH CHECK ADD  CONSTRAINT [FK_", worksheetName, "_", varName, "_", varName, "Id] FOREIGN KEY([", varName, "Id]) REFERENCES[dbo].[", varName, "]([Id]) ON DELETE CASCADE\nGO\n[REFERENCES]"));
                return SQLWrapperGenerationResult;
            }

            return SQLWrapperGenerationResult;
        }
        private string GetSQLColType(ExcelRange cell)
        {
            if (cell.Style.Numberformat.Format == "0")
                return "INT";

            return "NVARCHAR(200)";
        }
        #endregion

        public void GenerateClassSchema()
        {
            // get templates
            string resultFileName = Path.Combine(Path.GetTempPath(), "Results.cs");
            string newNamespaceTemplate = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Templates\\NewNamespaceTemplate.txt");
            string newClassTemplate = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\Templates\\NewClassTemplate.txt");

            //open excel data file, loop through each worksheet and read tables
            using (ExcelPackage package = new ExcelPackage(new FileInfo(DocPath)))
            {
                int colCount = 0;
                int rowCount = 0;

                string classGenerationResults = String.Empty;
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    string worksheetClassGenerationResult = newClassTemplate.Replace("[CLASS_NAME]", worksheet.Name);

                    //get Column and Row Count. Do not consider empty columns
                    colCount = worksheet.Dimension.End.Column;
                    rowCount = worksheet.Dimension.End.Row;
                    for (int row = 1; row <= 1; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var currentCell = worksheet.Cells[row, col];
                            worksheetClassGenerationResult = UpdateClassProp(currentCell, worksheetClassGenerationResult);
                        }

                        //remove placeholder
                        worksheetClassGenerationResult = worksheetClassGenerationResult.Replace("        [PROP_NAME]\r\n", "");
                        classGenerationResults = string.Concat(classGenerationResults, worksheetClassGenerationResult);
                    }
                }

                //finally save file 
                string result = newNamespaceTemplate.Replace("    [CLASSES]", classGenerationResults);
                result = result.Replace("[NAMESPACE_NAME]", NamespaceToUse);
                File.WriteAllText(resultFileName, result);

                //open file for user using associated default application
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo(resultFileName)
                {
                    UseShellExecute = true
                };
                p.Start();
            }

        }

        #region Class Generation Helpers
        private string UpdateClassProp(ExcelRange cell, string worksheetClassGenerationResult)
        {
            string variableType = GetClassPropertyType(cell);

            if (cell.Text.Contains("FK_"))
            {
                //replace placeholder with new prop, then put placeholder on new line for next prop
                var varName = cell.Text.Replace("FK_", "");
                worksheetClassGenerationResult = worksheetClassGenerationResult.Replace("    [PROP_NAME]", String.Concat("    public ", variableType, " ", varName, "Id { get; set; }", "\n        [PROP_NAME]"));
                worksheetClassGenerationResult = worksheetClassGenerationResult.Replace("    [PROP_NAME]", String.Concat("    public ", varName, " ", varName, " { get; set; }", "\n        [PROP_NAME]"));
                return worksheetClassGenerationResult;
            }
            else
                return worksheetClassGenerationResult.Replace("    [PROP_NAME]", String.Concat("    public ", variableType, " ", cell.Text, " { get; set; }", "\n        [PROP_NAME]"));

            return string.Empty;
        }
        private string GetClassPropertyType(ExcelRange cell)
        {
            if (cell.Style.Numberformat.Format == "0")
                return "int";

            return "string";
        }
        #endregion 
    }
}
