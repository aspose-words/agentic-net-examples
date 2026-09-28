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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Set the table's preferred width to 15 centimeters.
        // 1 inch = 2.54 cm, 1 inch = 72 points.
        double widthPoints = 15.0 * 72.0 / 2.54; // Convert centimeters to points.
        table.PreferredWidth = PreferredWidth.FromPoints(widthPoints);

        // Save the document.
        string outputPath = "TablePreferredWidth.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
