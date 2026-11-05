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

        // Insert first cell with some text.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Insert second cell with some text.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Set left indent to 2 centimeters (approximately 56.6929 points).
        const float centimetersToPoints = 72f / 2.54f; // points per centimeter
        table.LeftIndent = 2 * centimetersToPoints;

        // Save the document.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
