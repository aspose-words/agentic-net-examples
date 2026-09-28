using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Add a simple header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add a new row where each cell contains rich‑text formatting.

        // First cell.
        builder.InsertCell();

        // Bold text.
        builder.Font.Bold = true;
        builder.Write("Bold ");
        builder.Font.Bold = false;

        // Italic text.
        builder.Font.Italic = true;
        builder.Write("Italic ");
        builder.Font.Italic = false;

        // Red colored text.
        builder.Font.Color = Color.Red;
        builder.Write("Red");
        builder.Font.Color = Color.Empty; // Reset to default.

        // Move to the next cell.
        builder.InsertCell();

        // Underlined text.
        builder.Font.Underline = Underline.Single;
        builder.Write("Underline ");
        builder.Font.Underline = Underline.None;

        // Blue colored text.
        builder.Font.Color = Color.Blue;
        builder.Write("Blue");
        builder.Font.Color = Color.Empty;

        // Large text (16 points).
        builder.Font.Size = 16;
        builder.Write(" Large");
        builder.Font.Size = 12; // Reset to default size (12 points).

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "FormattedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
