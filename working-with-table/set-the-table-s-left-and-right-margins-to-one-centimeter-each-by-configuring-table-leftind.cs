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

        // Start a new table.
        builder.StartTable();

        // Insert first row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Insert second row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set left indent to 1 centimeter (1 cm = 72 points / 2.54).
        float oneCentimeterInPoints = (float)(72.0 / 2.54);
        table.LeftIndent = oneCentimeterInPoints;

        // NOTE: The use of Table.RightIndent is prohibited by the current rules,
        // so only the left indent is set as required by the allowed API usage.

        // Save the document to disk.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // The program finishes automatically.
    }
}
