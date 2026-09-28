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

        // Insert a heading paragraph that we will later locate.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");

        // Add another paragraph so the document has more content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Some content before the table.");

        // Locate the heading paragraph (first paragraph with Heading1 style).
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

        // Build a simple 2x2 table manually.
        Table table = new Table(doc);

        // First row.
        Row row1 = new Row(doc);
        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        row1.Cells.Add(cell11);

        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        row1.Cells.Add(cell12);
        table.Rows.Add(row1);

        // Second row.
        Row row2 = new Row(doc);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        row2.Cells.Add(cell21);

        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        row2.Cells.Add(cell22);
        table.Rows.Add(row2);

        // Insert the table after the heading paragraph.
        headingParagraph.ParentNode.InsertAfter(table, headingParagraph);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Optionally, you could add further validation here.
    }
}
