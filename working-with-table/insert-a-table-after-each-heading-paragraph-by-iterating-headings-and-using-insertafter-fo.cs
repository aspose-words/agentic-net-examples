using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three heading paragraphs.
        for (int i = 1; i <= 3; i++)
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Heading {i}");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"Paragraph under heading {i}.");
        }

        // Collect all heading paragraphs.
        List<Paragraph> headingParagraphs = new List<Paragraph>();
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                headingParagraphs.Add(para);
            }
        }

        // Insert a simple 2x2 table after each heading.
        foreach (Paragraph heading in headingParagraphs)
        {
            Table table = CreateSampleTable(doc);
            // Insert the table after the heading paragraph.
            heading.ParentNode.InsertAfter(table, heading);
        }

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method to create a 2x2 table with sample text.
    private static Table CreateSampleTable(Document doc)
    {
        Table table = new Table(doc);

        // Optional: set a simple border for visibility.
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);

        // First row.
        Row row1 = new Row(doc);
        table.AppendChild(row1);
        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));
        row1.AppendChild(cell11);
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
        row1.AppendChild(cell12);

        // Second row.
        Row row2 = new Row(doc);
        table.AppendChild(row2);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
        row2.AppendChild(cell21);
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
        row2.AppendChild(cell22);

        return table;
    }
}
