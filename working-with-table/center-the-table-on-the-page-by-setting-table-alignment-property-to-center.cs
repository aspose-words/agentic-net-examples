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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Center the table on the page.
            table.Alignment = TableAlignment.Center;

            // Define output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CenteredTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");

            // Optionally, you could add further processing here.
        }
    }
}
