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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple paragraph.
        builder.Writeln("This is a sample paragraph added to the document.");

        // Insert a table with two rows and two columns.
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

        // Insert a bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Bullet item 1");
        builder.Writeln("Bullet item 2");
        builder.Writeln("Bullet item 3");
        // Stop list formatting.
        builder.ListFormat.RemoveNumbers();

        // Prepare output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SampleDocument.odt");

        // Save the document as ODT using OdtSaveOptions.
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        doc.Save(outputPath, saveOptions);
    }
}
