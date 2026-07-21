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

        // Start a table.
        Table table = builder.StartTable();

        // First row – two cells with sample text.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row – two cells with sample text.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set the height rule of the second row to Auto (no explicit Height value).
        Row secondRow = table.Rows[1];
        secondRow.RowFormat.HeightRule = HeightRule.Auto;

        // Save the document to the local file system.
        doc.Save("RowHeightAuto.docx");
    }
}
