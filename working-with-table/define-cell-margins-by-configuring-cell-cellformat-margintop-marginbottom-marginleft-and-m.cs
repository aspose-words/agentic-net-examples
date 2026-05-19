using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMarginsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // Insert the first cell and set its margins (implemented as padding).
            builder.InsertCell();
            builder.CellFormat.TopPadding = 10;    // MarginTop
            builder.CellFormat.BottomPadding = 15; // MarginBottom
            builder.CellFormat.LeftPadding = 20;   // MarginLeft
            builder.CellFormat.RightPadding = 25;  // MarginRight
            builder.Write("Cell with custom margins.");

            // Insert a second cell with default margins.
            builder.InsertCell();
            builder.Write("Second cell.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellMargins.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The document was not saved successfully.");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
