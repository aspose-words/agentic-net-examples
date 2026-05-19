using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a few cells.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable right‑to‑left direction for the entire table.
        table.Bidi = true;

        // Simple validation.
        if (!table.Bidi)
            throw new InvalidOperationException("Table direction was not set to right‑to‑left.");

        // Save the document.
        const string outputFile = "TableRightToLeft.docx";
        doc.Save(outputFile);
    }
}
