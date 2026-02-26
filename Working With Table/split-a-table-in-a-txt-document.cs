using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // <-- added namespace for Table, Row, etc.

class SplitTableToTxt
{
    static void Main()
    {
        // Load the source document (must contain at least one table).
        Document srcDoc = new Document("TableSource.docx");

        // Create a new empty document that will hold the split tables.
        Document dstDoc = new Document();
        // Ensure the document has a section and a body.
        dstDoc.EnsureMinimum();

        // Iterate over all tables in the source document.
        foreach (Table srcTable in srcDoc.GetChildNodes(NodeType.Table, true))
        {
            // For each row in the source table create a new table that contains only that row.
            foreach (Row srcRow in srcTable.Rows)
            {
                // Clone the row (deep copy) so that it can be moved to a new table.
                Row clonedRow = (Row)srcRow.Clone(true);

                // Create a new table and add the cloned row.
                Table newTable = new Table(dstDoc);
                newTable.AppendChild(clonedRow);

                // Append the new table to the destination document body.
                dstDoc.FirstSection.Body.AppendChild(newTable);

                // Add an empty paragraph after each table to separate them in the text output.
                dstDoc.FirstSection.Body.AppendChild(new Paragraph(dstDoc));
            }
        }

        // Configure TXT save options to preserve the visual layout of tables.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true // Apply whitespace padding so the table shape is kept.
        };

        // Save the resulting document as plain‑text.
        dstDoc.Save("SplitTables.txt", txtOptions);
    }
}
