using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple paragraph.
        builder.Writeln("This is a paragraph before the table and list.");

        // Build a 2x2 table.
        Table table = builder.StartTable();

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

        // Finish the table.
        builder.EndTable();

        // Add a blank line between the table and the list.
        builder.Writeln();

        // Create a bulleted list with three items.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");
        // Reset list formatting so subsequent text is normal.
        builder.ListFormat.RemoveNumbers();

        // Prepare ODT save options (optional: set measurement unit).
        OdtSaveOptions saveOptions = new OdtSaveOptions
        {
            MeasureUnit = OdtSaveMeasureUnit.Centimeters
        };

        // Save the document as an ODT file using the specified options.
        doc.Save("OutputDocument.odt", saveOptions);
    }
}
