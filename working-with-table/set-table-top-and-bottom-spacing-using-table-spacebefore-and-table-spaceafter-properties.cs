using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableSpacingExample
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
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Set the spacing before and after the table (in points).
            // Use DistanceTop for space above the table and DistanceBottom for space below.
            table.DistanceTop = 20;    // 20 points above the table.
            table.DistanceBottom = 30; // 30 points below the table.

            // Save the document to disk.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableSpacing.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved successfully.");
        }
    }
}
