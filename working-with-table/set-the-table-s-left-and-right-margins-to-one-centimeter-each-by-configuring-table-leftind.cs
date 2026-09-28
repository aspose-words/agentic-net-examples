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

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set left indent to 1 centimeter (approximately 28.35 points).
        table.LeftIndent = 28.35f;

        // Note: Aspose.Words does not provide a Table.RightIndent property.
        // Right margin can be controlled via other layout settings if needed.

        // Save the document to the local file system.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output file was not created.");
        }
    }
}
