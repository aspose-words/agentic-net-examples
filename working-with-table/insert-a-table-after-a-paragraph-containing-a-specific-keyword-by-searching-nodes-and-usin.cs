using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph before the keyword.");
        builder.Writeln("Please INSERT_HERE the table below."); // Keyword paragraph.
        builder.Writeln("Another paragraph after the keyword.");

        // Search for the paragraph that contains the specific keyword.
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            string text = para.ToString(SaveFormat.Text);
            if (text.Contains("INSERT_HERE"))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Keyword paragraph not found.");

        // Build a simple 2x2 table.
        Table table = new Table(doc);
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // First row.
        Row row1 = new Row(doc);
        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        row1.AppendChild(cell11);

        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        row1.AppendChild(cell12);
        table.AppendChild(row1);

        // Second row.
        Row row2 = new Row(doc);
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        row2.AppendChild(cell21);

        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        row2.AppendChild(cell22);
        table.AppendChild(row2);

        // Insert the table after the paragraph containing the keyword.
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
