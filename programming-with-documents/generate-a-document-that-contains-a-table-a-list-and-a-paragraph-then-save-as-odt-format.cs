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

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple paragraph.
        builder.Writeln("This is a sample paragraph.");

        // Insert a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First list item");
        builder.Writeln("Second list item");
        builder.Writeln("Third list item");
        builder.ListFormat.RemoveNumbers();

        // Insert a 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Prepare ODT save options (default options are sufficient).
        OdtSaveOptions saveOptions = new OdtSaveOptions();

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.odt");

        // Save the document as ODT.
        doc.Save(outputPath, saveOptions);
    }
}
