using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SplitTableInPdf
{
    static void Main()
    {
        // Load the PDF document (Aspose.Words can import PDF files).
        Document doc = new Document("Input.pdf");

        // Ensure the document has at least one table.
        if (doc.FirstSection.Body.Tables.Count == 0)
        {
            Console.WriteLine("No tables found in the document.");
            return;
        }

        // Get the first table to split.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define after which row the table should be split (zero‑based index).
        // For example, split after the second row (i.e., rows 0 and 1 stay in the original table).
        int splitAfterRowIndex = 1;

        // Validate the split index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Create a new table that will hold the rows after the split point.
        Table newTable = new Table(doc);

        // Copy basic formatting from the original table to the new one.
        newTable.AllowAutoFit = originalTable.AllowAutoFit;
        newTable.Alignment = originalTable.Alignment;
        newTable.PreferredWidth = originalTable.PreferredWidth;
        newTable.Style = originalTable.Style;
        newTable.StyleIdentifier = originalTable.StyleIdentifier;
        newTable.StyleName = originalTable.StyleName;

        // Move rows after the split point from the original table to the new table.
        // Rows are removed from the original table as they are appended to the new table.
        // The loop continues until only rows up to splitAfterRowIndex remain.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // The row to move is always the one immediately after the split index.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];
            originalTable.RemoveChild(rowToMove);
            newTable.AppendChild(rowToMove);
        }

        // Insert the new table directly after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document as PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Optional: set page layout if desired.
        // saveOptions.PageLayout = PdfPageLayout.OneColumn;

        doc.Save("Output.pdf", saveOptions);
    }
}
