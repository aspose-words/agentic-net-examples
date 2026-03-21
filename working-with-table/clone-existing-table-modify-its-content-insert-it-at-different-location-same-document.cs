using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document with some content and a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("First paragraph.");

        // Create a simple table with two cells.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Original Cell 1");
        builder.InsertCell();
        builder.Writeln("Original Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Second paragraph (the insertion point will be after this).
        builder.Writeln("Second paragraph.");

        // Find the first table in the document.
        Table sourceTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (sourceTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Perform a deep clone of the table, including all rows, cells and runs.
        Table clonedTable = (Table)sourceTable.Clone(true);

        // ----- Modify the cloned table -----
        // Example: change the text in the first cell of the first row.
        Paragraph firstParagraph = clonedTable.FirstRow.FirstCell.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.Runs.Clear();
            firstParagraph.AppendChild(new Run(doc, "Modified content"));
        }

        // ----- Insert the cloned table at a new location -----
        // Example: insert after the second paragraph in the document.
        Paragraph insertionPoint = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        if (insertionPoint != null && insertionPoint.ParentNode != null)
        {
            insertionPoint.ParentNode.InsertAfter(clonedTable, insertionPoint);
        }
        else
        {
            // If the expected paragraph is not found, append the table at the end.
            doc.FirstSection.Body.AppendChild(clonedTable);
        }

        // Save the updated document.
        doc.Save("Result.docx");
        Console.WriteLine("Document saved as Result.docx");
    }
}
