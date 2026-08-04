using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAutoFit
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row with longer text to demonstrate auto‑fit.
            builder.InsertCell();
            builder.Write("This is a very long piece of text that should cause the column to expand.");
            builder.InsertCell();
            builder.Write("Short");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Enable auto‑fit to contents so columns resize based on their text.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "AutoFitTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // Inform the user where the file was saved.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
