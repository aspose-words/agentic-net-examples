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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table (first table in the document).
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set the left indent to two centimeters.
        // 1 centimeter = 28.3464567 points.
        double twoCentimetersInPoints = 2 * 28.3464567;
        table.LeftIndent = twoCentimetersInPoints;

        // Save the document to a file.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
