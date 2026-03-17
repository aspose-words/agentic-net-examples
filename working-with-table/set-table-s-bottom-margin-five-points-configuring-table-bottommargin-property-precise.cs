using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetTableBottomMargin
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell with some text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell content.");
        builder.EndTable();

        // Configure the distance between the table bottom and surrounding text to 5 points.
        // This property controls the table's bottom margin.
        table.DistanceBottom = 5.0;

        // Save the document to the desired location.
        doc.Save("Table.BottomMargin.docx");
    }
}
