using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Paragraph ----------
        builder.Writeln("This is a sample paragraph added to the document.");

        // ---------- List ----------
        // Start a bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");
        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Add an empty line between list and table for readability.
        builder.Writeln();

        // ---------- Table ----------
        // Start a 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // ---------- Save as ODT ----------
        // Ensure the output directory exists.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.odt");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Use OdtSaveOptions to specify ODT format.
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        doc.Save(outputPath, saveOptions);

        // Optional: indicate completion (no interactive input).
        Console.WriteLine("Document created and saved as ODT at: " + outputPath);
    }
}
