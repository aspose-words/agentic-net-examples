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

        // Start a table and add a few rows with two columns each.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
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

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable the first‑column conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Retrieve the style object to modify its conditional style for the first column.
        TableStyle tableStyle = (TableStyle)doc.Styles[StyleIdentifier.MediumShading1Accent1];
        // Make the text in the first column bold.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleFirstColumnBold.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
