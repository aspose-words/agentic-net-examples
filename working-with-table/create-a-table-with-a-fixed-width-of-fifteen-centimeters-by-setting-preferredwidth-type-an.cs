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

        // Start a new table.
        Table table = builder.StartTable();

        // Insert a single cell so that the table has at least one row.
        builder.InsertCell();
        builder.Write("Fixed width table cell.");

        // End the first (and only) row.
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Convert 15 centimeters to points (1 inch = 2.54 cm, 1 point = 1/72 inch).
        // points = (cm / 2.54) * 72
        double points = (15.0 / 2.54) * 72.0;

        // Set the table's preferred width to the calculated points value.
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document to the local file system.
        doc.Save("TableFixedWidth.docx");
    }
}
