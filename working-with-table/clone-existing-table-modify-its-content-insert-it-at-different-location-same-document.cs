using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Source.docx");

        // Find the table that we want to clone (first table in the document).
        Table sourceTable = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Perform a deep clone of the table, including all rows, cells and runs.
        Table clonedTable = (Table)sourceTable.Clone(true);

        // ----- Modify the cloned table -----
        // Example: change the text in the first cell of the first row.
        Paragraph firstParagraph = clonedTable.FirstRow.FirstCell.FirstParagraph;

        // Remove any existing runs in that paragraph.
        firstParagraph.Runs.Clear();

        // Add a new run with the desired text.
        firstParagraph.AppendChild(new Run(doc, "Modified content"));

        // ----- Insert the cloned table at a new location -----
        // Example: insert after the second paragraph in the document.
        Paragraph insertionPoint = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);

        // Insert the cloned table after the chosen paragraph.
        insertionPoint.ParentNode.InsertAfter(clonedTable, insertionPoint);

        // Save the updated document.
        doc.Save("Result.docx");
    }
}
