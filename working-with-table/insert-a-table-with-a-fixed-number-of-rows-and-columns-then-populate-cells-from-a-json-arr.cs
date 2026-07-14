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
            // Each inner array is a row; each string is a cell value.
            string json = @"
            [
                [""Row 1, Col 1"", ""Row 1, Col 2"", ""Row 1, Col 3""],
                [""Row 2, Col 1"", ""Row 2, Col 2"", ""Row 2, Col 3""],
                [""Row 3, Col 1"", ""Row 3, Col 2"", ""Row 3, Col 3""]
            ]";

            // Deserialize JSON into a list of rows, each row being a list of cell strings.
            List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // Populate the table using the data from the JSON array.
            foreach (List<string> row in tableData)
            {
                foreach (string cellText in row)
                {
                    builder.InsertCell();
                    builder.Write(cellText);
                }
                // End the current row and start a new one.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Define output path (in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromJson.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Failed to create the output document.");

            // The program ends here without waiting for user input.
        }
    }
}
