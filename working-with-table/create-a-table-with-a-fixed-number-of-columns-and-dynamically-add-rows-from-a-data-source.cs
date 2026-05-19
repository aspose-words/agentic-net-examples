using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a simple data source: each inner list represents a row of values.
            List<string[]> dataRows = new List<string[]>
            {
                new[] { "Alice", "Engineering", "85" },
                new[] { "Bob", "Marketing", "78" },
                new[] { "Charlie", "Finance", "92" }
            };

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with a fixed number of columns (3 in this example).
            Table table = builder.StartTable();

            // Add a header row.
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Department");
            builder.InsertCell();
            builder.Write("Score");
            builder.EndRow();

            // Dynamically add rows from the data source.
            foreach (string[] rowData in dataRows)
            {
                // Ensure each row has exactly three cells.
                for (int i = 0; i < 3; i++)
                {
                    builder.InsertCell();
                    builder.Write(rowData[i]);
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Define the output file path (saved in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // The program ends here; no user interaction is required.
        }
    }
}
