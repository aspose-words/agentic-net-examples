using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableAfterHeading
{
    static void Main()
    {
        // Create a new document and add a heading paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");

        // Add a paragraph after the heading to show where the table will be inserted.
        builder.Writeln("Paragraph before table.");

        // Locate the heading paragraph by its style.
        Paragraph heading = doc.GetChildNodes(NodeType.Paragraph, true)
                              .OfType<Paragraph>()
                              .First(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1);

        // Create a 2x2 table.
        Table table = new Table(doc);
        table.EnsureMinimum(); // Guarantees at least one row, cell, and paragraph.

        // First row cells.
        Row firstRow = table.FirstRow;
        Cell cell11 = firstRow.FirstCell;
        cell11.FirstParagraph.AppendChild(new Run(doc, "R1C1"));
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "R1C2"));
        firstRow.AppendChild(cell12);

        // Second row cells.
        Row secondRow = new Row(doc);
        table.AppendChild(secondRow);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "R2C1"));
        secondRow.AppendChild(cell21);
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "R2C2"));
        secondRow.AppendChild(cell22);

        // Insert the table immediately after the heading paragraph.
        heading.ParentNode.InsertAfter(table, heading);

        // Save the document.
        doc.Save("Output.docx");
    }
}
