using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBorders
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 2x2 table.
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
            builder.EndTable(); // Ends the table.

            // Apply a custom border thickness of 2 points to all sides of the table.
            // Use the SetBorders method because Table.Borders is not available.
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBorders.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
