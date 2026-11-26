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

        // Add a new row with rich‑text formatted runs in each cell.
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

        // Underlined text.
        builder.Font.Underline = Underline.Single;
        builder.Write("Underline");
        builder.Font.Underline = Underline.None;

        // End the paragraph in the first cell.
        builder.Writeln();

        // Second cell.
        builder.InsertCell();

        // Red text.
        builder.Font.Color = Color.Red;
        builder.Write("Red ");

        // Blue text.
        builder.Font.Color = Color.Blue;
        builder.Write("Blue ");

        // Green text.
        builder.Font.Color = Color.Green;
        builder.Write("Green");

        // Reset color and end paragraph.
        builder.Font.Color = Color.Empty;
        builder.Writeln();

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
