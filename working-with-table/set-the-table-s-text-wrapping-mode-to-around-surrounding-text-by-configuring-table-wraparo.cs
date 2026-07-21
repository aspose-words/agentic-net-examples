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

        // Set a preferred width so the floating table is visible.
        table.PreferredWidth = PreferredWidth.FromPoints(200);

        // Configure the table to wrap text around it.
        table.TextWrapping = TextWrapping.Around;
        // Optional: define the distance between the table and surrounding text.
        table.AbsoluteHorizontalDistance = 20; // points
        table.AbsoluteVerticalDistance = 10;   // points

        // Add a paragraph after the table to demonstrate the wrapping effect.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapAround.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
