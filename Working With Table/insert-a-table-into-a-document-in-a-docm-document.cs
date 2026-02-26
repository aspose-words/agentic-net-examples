using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsert
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing DOCM template.
            string dataDir = @"C:\Data\";
            string inputPath = System.IO.Path.Combine(dataDir, "Template.docm");
            string outputPath = System.IO.Path.Combine(dataDir, "Output.docm");

            // Load the DOCM document.
            Document doc = new Document(inputPath);

            // Create a DocumentBuilder attached to the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple 2‑cell table at the current cursor position.
            Table table = builder.StartTable();          // Starts a new table and returns the Table node.
            builder.InsertCell();                        // First cell.
            builder.Write("First cell");                 // Add text to the first cell.
            builder.InsertCell();                        // Second cell.
            builder.Write("Second cell");                // Add text to the second cell.
            builder.EndRow();                            // End the first row.
            builder.EndTable();                          // Finish the table.

            // Optionally set table title/description (useful for DOCX/ DOCM compliance).
            table.Title = "Sample Table";
            table.Description = "A table inserted programmatically into a DOCM file.";

            // Save the modified document as DOCM.
            doc.Save(outputPath, SaveFormat.Docm);
        }
    }
}
