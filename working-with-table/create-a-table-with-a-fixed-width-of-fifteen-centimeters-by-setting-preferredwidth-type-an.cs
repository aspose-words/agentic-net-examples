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

        // Build a simple 2‑column table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set the table width to a fixed 15 centimeters.
        // 1 inch = 72 points, 1 cm = 72 / 2.54 points.
        double centimeters = 15;
        double points = centimeters * 72.0 / 2.54;
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document.
        string fileName = "TableFixedWidth.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output file was not created.");
    }
}
