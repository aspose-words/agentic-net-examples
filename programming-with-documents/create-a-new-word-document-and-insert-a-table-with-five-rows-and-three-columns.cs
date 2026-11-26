using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the table.
            builder.StartTable();

            // Insert 5 rows and 3 columns.
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    // Insert a new cell and add placeholder text.
                    builder.InsertCell();
                    builder.Writeln($"Row {row + 1}, Col {col + 1}");
                }
                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "TableDocument.docx";
            doc.Save(outputPath);

            // Confirm that the file was saved.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
            }
        }
    }
}
