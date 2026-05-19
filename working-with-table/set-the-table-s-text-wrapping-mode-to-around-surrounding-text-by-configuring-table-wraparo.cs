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

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Set a preferred width so that surrounding text has space to flow.
        table.PreferredWidth = PreferredWidth.FromPoints(200);

        // Configure the table to wrap text around it.
        table.TextWrapping = TextWrapping.Around;
        // Optional: distance from surrounding text (in points).
        table.AbsoluteHorizontalDistance = 10;
        table.AbsoluteVerticalDistance = 10;

        // Add a paragraph after the table to demonstrate wrapping.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.WrapAround.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Output file was not created.");
    }
}
