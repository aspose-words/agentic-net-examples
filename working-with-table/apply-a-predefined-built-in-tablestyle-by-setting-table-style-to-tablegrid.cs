using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add two rows with two cells each.
            Table table = builder.StartTable();

            // First row (header)
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row (data)
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply the built‑in "TableGrid" style to the table.
            // Use the StyleIdentifier property because Table.Style expects a Style object.
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Define the output file path (in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleExample.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to create the output document.");
            }
        }
    }
}
