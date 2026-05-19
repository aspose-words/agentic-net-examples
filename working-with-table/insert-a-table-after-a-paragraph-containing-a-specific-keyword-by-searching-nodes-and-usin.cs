using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class InsertTableAfterParagraph
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs. One of them contains the keyword we will search for.
        builder.Writeln("First paragraph.");
        builder.Writeln("Paragraph with the keyword: INSERT_HERE");
        builder.Writeln("Last paragraph.");

        // Search for the paragraph that contains the keyword.
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Contains("INSERT_HERE"))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Keyword paragraph not found.");

        // Create a table node and ensure it has at least one row and one cell.
        Table table = new Table(doc);
        table.EnsureMinimum();

        // Populate the first row with two cells.
        Cell cell1 = table.FirstRow.FirstCell;
        cell1.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));

        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        table.FirstRow.AppendChild(cell2);

        // Add a second row with two cells.
        Row secondRow = new Row(doc);
        table.AppendChild(secondRow);

        Cell cell3 = new Cell(doc);
        cell3.AppendChild(new Paragraph(doc));
        cell3.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        secondRow.AppendChild(cell3);

        Cell cell4 = new Cell(doc);
        cell4.AppendChild(new Paragraph(doc));
        cell4.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        secondRow.AppendChild(cell4);

        // Insert the table after the target paragraph.
        // The parent of the paragraph (a Body node) performs the insertion.
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the resulting document.
        string outputPath = "OutputTableAfterParagraph.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new IOException($"Failed to create the output file: {outputPath}");
    }
}
