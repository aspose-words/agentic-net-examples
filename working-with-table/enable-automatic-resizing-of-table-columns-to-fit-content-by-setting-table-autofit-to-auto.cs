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

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity (kg)");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable automatic column resizing to fit the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to a file.
        doc.Save("AutoFitTable.docx");
    }
}
