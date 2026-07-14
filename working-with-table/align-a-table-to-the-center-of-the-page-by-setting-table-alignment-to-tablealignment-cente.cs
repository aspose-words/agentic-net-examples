using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Align the table to the center of the page.
            table.Alignment = TableAlignment.Center;

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Failed to create the output document.");

            // Indicate successful completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
