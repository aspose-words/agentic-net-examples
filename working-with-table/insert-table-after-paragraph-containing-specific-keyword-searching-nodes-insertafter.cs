using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableAfterKeyword
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // The keyword to search for.
        const string keyword = "PLACEHOLDER";

        // Find the first paragraph that contains the keyword.
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Contains(keyword))
            {
                targetParagraph = para;
                break;
            }
        }

        // If the paragraph was not found, exit.
        if (targetParagraph == null)
        {
            Console.WriteLine($"Keyword \"{keyword}\" not found.");
            return;
        }

        // Create a simple 2x2 table.
        Table table = new Table(doc);
        // Ensure the table has at least one row and cell.
        table.EnsureMinimum();

        // Populate the first row.
        Row row1 = table.FirstRow;
        Cell cell11 = row1.FirstCell;
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
        row1.AppendChild(cell12);

        // Add a second row.
        Row row2 = new Row(doc);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
        row2.AppendChild(cell21);
        row2.AppendChild(cell22);
        table.AppendChild(row2);

        // Insert the table immediately after the found paragraph.
        // The paragraph's parent is a CompositeNode (e.g., Body).
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
