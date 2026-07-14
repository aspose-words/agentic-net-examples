using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableWrapExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a couple of cells with sample text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            // Set the table's text wrapping style to "Around" (square-like wrapping).
            table.TextWrapping = TextWrapping.Around;

            // Optional: set distances so the text appears around the table.
            table.AbsoluteHorizontalDistance = 20;
            table.AbsoluteVerticalDistance = 10;

            // Define an output path in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapSquare.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
