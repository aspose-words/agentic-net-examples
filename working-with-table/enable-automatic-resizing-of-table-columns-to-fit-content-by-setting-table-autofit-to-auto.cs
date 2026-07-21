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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert first row with two cells containing long text to demonstrate autofit.
            builder.InsertCell();
            builder.Write("This is a very long piece of text that should cause the column to expand when autofit is applied.");
            builder.InsertCell();
            builder.Write("Short text");
            builder.EndRow();

            // Insert second row.
            builder.InsertCell();
            builder.Write("Another long text entry that will test the auto resizing behavior of the table columns.");
            builder.InsertCell();
            builder.Write("Another short");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Enable automatic resizing of columns to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Define output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
            {
                throw new Exception($"Failed to create the output file at '{outputPath}'.");
            }

            // Optionally, inform that the process completed successfully.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
