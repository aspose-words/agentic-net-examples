using System;
using Aspose.Words;
using Aspose.Words.Tables;

class CloneTableExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add the first paragraph.
        builder.Writeln("First paragraph.");

        // Create a simple table.
        Table originalTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a second paragraph that will serve as the reference point.
        builder.Writeln("Second paragraph.");

        // Perform a deep clone of the table (including all rows, cells, and contents).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Locate the reference paragraph (the second paragraph in the document).
        Paragraph referenceParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        if (referenceParagraph == null)
        {
            throw new InvalidOperationException("Reference paragraph not found.");
        }

        // Insert the cloned table immediately after the reference paragraph.
        referenceParagraph.ParentNode.InsertAfter(clonedTable, referenceParagraph);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
