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

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity (kg)");
        builder.EndRow();

        // ----- Data rows -----
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style that supports conditional formatting.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable the first‑row (header) conditional formatting of the style.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Resize the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the local file system.
        const string outputFile = "TableWithHeaderStyle.docx";
        doc.Save(outputFile);
    }
}
