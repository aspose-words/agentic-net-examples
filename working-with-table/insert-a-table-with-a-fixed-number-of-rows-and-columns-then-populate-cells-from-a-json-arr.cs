using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // JSON array representing rows and columns of the table.
            // Each inner array is a row, each string is a cell value.
            string json = @"
            [
                [""Header 1"", ""Header 2"", ""Header 3""],
                [""Row1 Col1"", ""Row1 Col2"", ""Row1 Col3""],
                [""Row2 Col1"", ""Row2 Col2"", ""Row2 Col3""]
            ]";

            // Deserialize JSON into a list of rows, each row is a list of cell strings.
            List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

            if (tableData == null || tableData.Count == 0)
                throw new InvalidOperationException("JSON does not contain any table data.");

            // Determine the number of rows and columns.
            int rowCount = tableData.Count;
            int columnCount = tableData[0].Count;

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // Populate the table cells from the JSON data.
            for (int i = 0; i < rowCount; i++)
            {
                // Ensure each row has the expected number of columns.
                if (tableData[i].Count != columnCount)
                    throw new InvalidOperationException($"Row {i} does not have the expected {columnCount} columns.");

                for (int j = 0; j < columnCount; j++)
                {
                    // Insert a new cell and write its text.
                    builder.InsertCell();
                    builder.Write(tableData[i][j]);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Define output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
        }
    }
}
