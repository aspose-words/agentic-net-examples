using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOCX that contains content controls (structured document tags).
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // Example 1: Update a plain‑text content control.
        // ------------------------------------------------------------
        // Find the first StructuredDocumentTag in the document.
        StructuredDocumentTag plainTextSdt = doc.GetChild(NodeType.StructuredDocumentTag, 0, true) as StructuredDocumentTag;

        if (plainTextSdt != null && plainTextSdt.SdtType == SdtType.PlainText)
        {
            // Remove any existing child nodes (the old text).
            plainTextSdt.RemoveAllChildren();

            // Insert the new text that should appear inside the control.
            plainTextSdt.AppendChild(new Run(doc, "New content for the plain‑text control"));
        }

        // ------------------------------------------------------------
        // Example 2: Update a repeating‑section content control.
        // ------------------------------------------------------------
        // Find the second StructuredDocumentTag (index 1) – assumed to be a repeating section.
        StructuredDocumentTag repeatingSdt = doc.GetChild(NodeType.StructuredDocumentTag, 1, true) as StructuredDocumentTag;

        if (repeatingSdt != null && repeatingSdt.SdtType == SdtType.RepeatingSection)
        {
            // The repeating section usually contains a table that represents the repeated rows.
            Table table = repeatingSdt.GetChild(NodeType.Table, 0, true) as Table;

            if (table != null)
            {
                // Clone the last row to create a new row with the same formatting.
                Row newRow = table.LastRow.Clone(true) as Row;

                // Clear existing cell contents (optional, depending on template).
                foreach (Cell cell in newRow.Cells)
                    cell.RemoveAllChildren();

                // Populate cells with new data.
                newRow.Cells[0].FirstParagraph.AppendChild(new Run(doc, "Item 1"));
                newRow.Cells[1].FirstParagraph.AppendChild(new Run(doc, "Value 1"));
                newRow.Cells[2].FirstParagraph.AppendChild(new Run(doc, "Description 1"));

                // Add the new row to the table.
                table.Rows.Add(newRow);
            }
        }

        // ------------------------------------------------------------
        // Save the modified document.
        // ------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
