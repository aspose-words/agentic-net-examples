using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetCellTextDirection
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // -------------------------------------------------------------------
        // NOTE: The CellFormat.TextDirection property exists only in Aspose.Words
        // versions 23.5 and later. If you are using an older version the property
        // is not available, which causes the CS1061 compile error.
        // -------------------------------------------------------------------
        // Work‑around for older versions: set the paragraph inside the cell to
        // right‑to‑left (Bidi) and align the text to the right.
        builder.ParagraphFormat.Bidi = true;               // Enable right‑to‑left layout
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        // Write Arabic text into the cell.
        builder.Write("مرحبا بالعالم");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("CellTextDirectionRTL.docx");
    }
}
