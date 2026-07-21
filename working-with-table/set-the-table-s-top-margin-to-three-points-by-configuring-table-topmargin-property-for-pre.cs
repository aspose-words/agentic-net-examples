using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for constructing the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Write("Sample cell content.");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set the distance between the table top and surrounding text to 3 points.
        // This property controls the vertical placement of the table relative to surrounding paragraphs.
        table.DistanceTop = 3.0;

        // Save the document to the local file system.
        doc.Save("TableTopMargin.docx");
    }
}
