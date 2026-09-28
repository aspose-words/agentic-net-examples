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

        // Add some paragraphs to the document.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Target Paragraph"); // This is the paragraph after which we will insert the table.
        builder.Writeln("Paragraph 3");

        // Locate the paragraph node with the exact text "Target Paragraph".
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim() == "Target Paragraph")
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Create a new table.
        Table table = new Table(doc);

        // First row.
        Row row1 = new Row(doc);
        table.AppendChild(row1);

        // First cell of first row.
        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        row1.AppendChild(cell11);

        // Second cell of first row.
        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        row1.AppendChild(cell12);

        // Second row.
        Row row2 = new Row(doc);
        table.AppendChild(row2);

        // First cell of second row.
        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        row2.AppendChild(cell21);

        // Second cell of second row.
        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        row2.AppendChild(cell22);

        // Insert the table after the target paragraph.
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Optionally, you could add further processing here.
    }
}
