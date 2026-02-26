using System;
using Aspose.Words;
using Aspose.Words.Tables; // <-- added namespace for Table
using Aspose.Words.Replacing;

class TableFindReplaceDemo
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Define find/replace options (case‑sensitive, whole‑word only).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Replace text in the whole table.
            table.Range.Replace("OldValue", "NewValue", options);

            // Example: replace text only in the last cell of the last row.
            if (table.LastRow != null && table.LastRow.LastCell != null)
            {
                table.LastRow.LastCell.Range.Replace("50", "20", options);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
