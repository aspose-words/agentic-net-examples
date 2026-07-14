using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading paragraph that we will later locate.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");

        // Add some regular text after the heading (optional).
        builder.ParagraphFormat.ClearFormatting();
        builder.Writeln("Some content before the table.");

        // Locate the heading paragraph by its style.
        Paragraph headingParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                para.GetText().Trim() == "Sample Heading")
            {
                headingParagraph = para;
                break;
            }
        }

        if (headingParagraph == null)
            throw new Exception("Heading paragraph not found.");

        // Create a new table (not yet attached to the document).
        Table table = new Table(doc);
        table.EnsureMinimum(); // Guarantees at least one row and one cell.

        // Populate the first row.
        Row firstRow = table.FirstRow;
        Cell firstCell = firstRow.FirstCell;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        Cell secondCell = new Cell(doc);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        firstRow.AppendChild(secondCell);

        // Add a second row.
        Row secondRow = new Row(doc);
        Cell thirdCell = new Cell(doc);
        thirdCell.AppendChild(new Paragraph(doc));
        thirdCell.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        secondRow.AppendChild(thirdCell);
        Cell fourthCell = new Cell(doc);
        fourthCell.AppendChild(new Paragraph(doc));
        fourthCell.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        secondRow.AppendChild(fourthCell);
        table.AppendChild(secondRow);

        // Insert the table after the heading paragraph.
        headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document.");
    }
}
