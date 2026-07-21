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
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Write("Fixed width table cell.");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Convert 15 centimeters to points (1 inch = 2.54 cm, 1 point = 1/72 inch).
        double points = 15.0 / 2.54 * 72.0;
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFixedWidth.docx");
        doc.Save(outputPath);
    }
}
