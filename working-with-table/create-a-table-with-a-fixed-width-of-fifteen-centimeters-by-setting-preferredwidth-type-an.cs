using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Write("Fixed width table cell.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Convert 15 centimeters to points (1 inch = 2.54 cm, 1 point = 1/72 inch).
        double points = 15.0 / 2.54 * 72.0;

        // Set the table's preferred width to the calculated points value.
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document to a file.
        const string outputPath = "TableFixedWidth.docx";
        doc.Save(outputPath);
    }
}
