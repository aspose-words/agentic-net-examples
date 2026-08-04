using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace TableTextDirectionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            // Set the cell's text orientation to vertical (top‑to‑bottom) for Asian scripts.
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("縦書きセル 1");

            // First row, second cell.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("縦書きセル 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("縦書きセル 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
            builder.Write("縦書きセル 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableTextDirection.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // Optional: inform that the process completed (no interactive prompts).
            Console.WriteLine("Document saved to: " + outputPath);
        }
    }
}
