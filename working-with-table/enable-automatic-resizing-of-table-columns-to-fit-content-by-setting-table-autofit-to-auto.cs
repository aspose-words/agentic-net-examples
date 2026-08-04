using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAutoFitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("This is a very long piece of text that should cause the column to expand automatically.");
            builder.InsertCell();
            builder.Write("Short");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Another long text entry to test auto‑fit behavior.");
            builder.InsertCell();
            builder.Write("Data");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply AutoFit to contents so columns resize based on their content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "AutoFitTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
