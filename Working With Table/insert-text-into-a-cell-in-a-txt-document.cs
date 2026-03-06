using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class InsertTextIntoCellInTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell.
        builder.StartTable();
        builder.InsertCell();               // Insert first (and only) cell.
        builder.Write("Hello, World!");     // Insert the desired text into the cell.
        builder.EndRow();                   // End the row.
        builder.EndTable();                 // End the table.

        // Configure TXT save options if needed (e.g., custom paragraph break).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Example: keep default settings or customize as required.
            // ParagraphBreak = "\r\n"
        };

        // Save the document as a plain‑text file.
        doc.Save("CellText.txt", txtOptions);
    }
}
