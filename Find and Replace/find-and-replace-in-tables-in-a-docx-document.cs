using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables; // Added namespace for Table class

class TableFindReplace
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options (case‑sensitive, whole‑word match).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };

        // Iterate through all tables in the document and replace text within each table.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Replace occurrences of "Carrots" with "Eggs" inside the current table.
            table.Range.Replace("Carrots", "Eggs", options);

            // Example: replace the value "50" with "20" in the last cell of the last row.
            if (table.LastRow?.LastCell != null)
                table.LastRow.LastCell.Range.Replace("50", "20", options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
