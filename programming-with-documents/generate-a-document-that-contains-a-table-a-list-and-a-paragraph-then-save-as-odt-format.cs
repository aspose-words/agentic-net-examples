using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple paragraph.
        builder.Writeln("This is a sample paragraph added to the document.");

        // Insert a 2x2 table.
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Insert a bullet list.
        builder.Writeln(); // Add a blank line before the list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Define output path.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.odt");

        // Save the document as ODT using OdtSaveOptions.
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        doc.Save(outputFile, saveOptions);
    }
}
