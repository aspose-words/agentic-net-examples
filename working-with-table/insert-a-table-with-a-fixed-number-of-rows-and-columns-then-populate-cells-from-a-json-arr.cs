using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableFromJson
{
    public class Program
    {
        public static void Main()
        {
            // JSON array representing rows and columns of the table.
            // Each inner array is a row, each string is a cell value.
            string json = @"[
                [""Row 1, Column 1"", ""Row 1, Column 2"", ""Row 1, Column 3""],
                [""Row 2, Column 1"", ""Row 2, Column 2"", ""Row 2, Column 3""],
                [""Row 3, Column 1"", ""Row 3, Column 2"", ""Row 3, Column 3""]
            ]";

            // Deserialize JSON into a list of rows, each row being a list of cell strings.
            List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

            if (tableData == null || tableData.Count == 0)
                throw new InvalidOperationException("JSON does not contain any table data.");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            builder.StartTable();

            // Iterate over rows.
            foreach (List<string> row in tableData)
            {
                // Iterate over cells in the current row.
                foreach (string cellText in row)
                {
                    builder.InsertCell();          // Insert a new cell.
                    builder.Write(cellText);       // Write the cell's text.
                }

                builder.EndRow(); // Finish the current row.
            }

            // Finish the table.
            builder.EndTable();

            // Define output path (in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
