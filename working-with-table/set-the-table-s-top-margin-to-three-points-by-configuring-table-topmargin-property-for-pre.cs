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

        // Start a table and add a single cell with some text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Set the distance between the top of the table and surrounding text to 3 points.
        // This property represents the table's top margin.
        table.DistanceTop = 3.0;

        // Save the document to the local file system.
        doc.Save("TableTopMargin.docx");
    }
}
