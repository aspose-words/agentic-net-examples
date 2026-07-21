using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class InsertTableAfterHeading
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading paragraph (Heading 1 style).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");

        // Add a normal paragraph after the heading (optional, just to have more content).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph following the heading.");

        // Locate the heading paragraph in the document.
        Paragraph headingParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                headingParagraph = para;
                break;
            }
        }

        if (headingParagraph == null)
            throw new InvalidOperationException("Heading paragraph not found.");

        // Create a new table node (not yet attached to the document).
        Table table = new Table(doc);

        // Ensure the table has at least one row, cell, and paragraph.
        table.EnsureMinimum();

        // Populate the first row with two cells.
        Row firstRow = table.FirstRow;
        Cell firstCell = firstRow.FirstCell;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        Cell secondCell = new Cell(doc);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        firstRow.AppendChild(secondCell);

        // Add a second row with two cells.
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
        // The heading's parent is the Body node; use InsertAfter on that parent.
        headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

        // Save the document.
        string outputPath = "OutputTableAfterHeading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new IOException($"Failed to create the output file: {outputPath}");
    }
}
