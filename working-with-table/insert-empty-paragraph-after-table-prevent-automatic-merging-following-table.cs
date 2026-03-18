using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertEmptyParagraphAfterTables
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate over all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Create an empty paragraph node.
            Paragraph emptyParagraph = new Paragraph(doc);

            // Insert the empty paragraph immediately after the current table.
            // The table's ParentNode is typically a Body node.
            table.ParentNode.InsertAfter(emptyParagraph, table);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
