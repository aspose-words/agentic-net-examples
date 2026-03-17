using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableAfterParagraph
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Find the paragraph after which we want to insert the table.
        // In this example we look for a paragraph that contains the exact text "Insert Table Here".
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim() == "Insert Table Here")
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
        {
            Console.WriteLine("Target paragraph not found.");
            return;
        }

        // Create a new table and ensure it has at least one row and one cell.
        Table table = new Table(doc);
        table.EnsureMinimum();

        // Populate the table with some sample data (2 rows x 2 columns).
        // First row.
        Row row1 = table.FirstRow;
        Cell cell11 = row1.FirstCell;
        cell11.FirstParagraph.AppendChild(new Run(doc, "R1C1"));
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "R1C2"));
        row1.AppendChild(cell12);

        // Second row.
        Row row2 = new Row(doc);
        table.AppendChild(row2);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "R2C1"));
        row2.AppendChild(cell21);
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "R2C2"));
        row2.AppendChild(cell22);

        // Insert the table after the target paragraph.
        // The paragraph's parent node is a Body (CompositeNode), so we use InsertAfter on it.
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
