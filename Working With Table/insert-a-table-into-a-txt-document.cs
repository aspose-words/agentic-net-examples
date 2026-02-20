using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add two rows with two cells each.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set a preferred width to help preserve layout in plain‑text output.
        table.PreferredWidth = PreferredWidth.FromPoints(200);

        // Configure save options to keep the table layout when exporting to TXT.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };

        // Save the document as a plain‑text file.
        doc.Save("TableInTxt.txt", saveOptions);
    }
}
