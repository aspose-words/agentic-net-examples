using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin building a table.
            Table table = builder.StartTable();

            // First row with two cells.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row with two cells.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a fixed width (300 points) and center the table on the page.
            table.PreferredWidth = PreferredWidth.FromPoints(300);
            table.Alignment = TableAlignment.Center;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "FixedWidthCenteredTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");
        }
    }
}
