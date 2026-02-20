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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optional: preserve the visual layout of the table when exporting to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };

        // Save the document as a TXT file.
        doc.Save("TableInTxtDocument.txt", saveOptions);
    }
}
