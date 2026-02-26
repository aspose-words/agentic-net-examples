using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class InsertTableIntoTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder associated with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, cell 1.");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, cell 2.");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, cell 1.");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, cell 2.");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Configure TXT save options to preserve table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };

        // Save the document as a plain‑text file.
        doc.Save("TableInTxt.txt", txtOptions);
    }
}
