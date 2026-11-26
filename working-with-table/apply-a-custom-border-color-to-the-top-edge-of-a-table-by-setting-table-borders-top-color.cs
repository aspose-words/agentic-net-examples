using System;
using System.Drawing;
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

        // Build a 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");

        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a custom red color to the top border of the table.
        // Use SetBorder because Table.Borders property is not available in this version.
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Red, true);

        // Define output path.
        string outputPath = "TableTopBorderColor.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
