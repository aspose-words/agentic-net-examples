using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoMhtmlCell
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Create a DocumentBuilder for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the desired cell.
        // Parameters: row index, column index, row span, column span (zero‑based indexes).
        // Adjust the indices to point to the target cell in your table.
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the desired text into the cell.
        builder.Write("Hello, Aspose.Words!");

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mhtml");
    }
}
