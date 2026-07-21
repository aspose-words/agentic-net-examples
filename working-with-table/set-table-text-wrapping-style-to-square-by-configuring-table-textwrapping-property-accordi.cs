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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Set a preferred width so the table does not occupy the whole line.
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            // Configure text wrapping around the table (square style).
            table.TextWrapping = TextWrapping.Around;
            // Optional: set distances from surrounding text.
            table.AbsoluteHorizontalDistance = 20;
            table.AbsoluteVerticalDistance = 10;

            // Add a paragraph after the table to demonstrate wrapping.
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Determine an output path and ensure the directory exists.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapText.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
