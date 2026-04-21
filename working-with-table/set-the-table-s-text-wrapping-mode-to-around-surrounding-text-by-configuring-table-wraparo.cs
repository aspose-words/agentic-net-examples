using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the table.
        builder.Writeln("This is some text before the table.");

        // Start a table and add two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Set the table width and configure text wrapping around it.
        table.PreferredWidth = PreferredWidth.FromPoints(200);
        table.TextWrapping = TextWrapping.Around;
        table.AbsoluteHorizontalDistance = 10; // space on the left/right of the table
        table.AbsoluteVerticalDistance = 10;   // space above/below the table

        // Add text after the table that will wrap around the floating table.
        builder.Writeln("This is some text after the table that should wrap around the table. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor " +
                        "incididunt ut labore et dolore magna aliqua.");

        // Save the document to a local file.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWrapAround.docx");
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
