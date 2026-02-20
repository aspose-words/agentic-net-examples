using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words;

class SplitTableInDotm
{
    static void Main()
    {
        // Load the DOTM template.
        Document doc = new Document(@"C:\Docs\Template.dotm");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Index of the row after which the split will occur (0‑based).
        // Rows with index < splitRowIndex stay in the original table,
        // rows with index >= splitRowIndex move to the new table.
        int splitRowIndex = 2;

        // Clone the original table to create a new table that will hold the second part.
        Table newTable = (Table)originalTable.Clone(true);

        // ----- Remove rows that belong to the second part from the original table -----
        for (int i = originalTable.Rows.Count - 1; i >= splitRowIndex; i--)
        {
            originalTable.Rows[i].Remove();
        }

        // ----- Remove rows that belong to the first part from the new table -----
        // Keep rows with index >= splitRowIndex, delete the ones before the split.
        for (int i = splitRowIndex - 1; i >= 0; i--)
        {
            newTable.Rows[i].Remove();
        }

        // Insert the new table immediately after the original table in the document tree.
        // InsertAfter is defined on CompositeNode, so cast the parent node.
        CompositeNode parent = (CompositeNode)originalTable.ParentNode;
        parent.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save(@"C:\Docs\Template_Split.dotm");
    }
}
