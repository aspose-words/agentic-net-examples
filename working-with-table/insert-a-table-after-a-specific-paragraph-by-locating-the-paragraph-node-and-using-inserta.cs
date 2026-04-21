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

        // Add some paragraphs – the second one will be the reference point.
        builder.Writeln("First paragraph.");
        builder.Writeln("Target paragraph."); // <-- we will insert the table after this paragraph.
        builder.Writeln("Third paragraph.");

        // Locate the paragraph node that contains the exact text "Target paragraph."
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim() == "Target paragraph.")
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Build a simple 2x2 table manually (without using DocumentBuilder.InsertTable).
        Table table = new Table(doc);

        // First row.
        Row row1 = new Row(doc);
        table.AppendChild(row1);

        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1"));
        row1.AppendChild(cell11);

        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 2"));
        row1.AppendChild(cell12);

        // Second row.
        Row row2 = new Row(doc);
        table.AppendChild(row2);

        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 3"));
        row2.AppendChild(cell21);

        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 4"));
        row2.AppendChild(cell22);

        // Insert the table after the located paragraph.
        // The paragraph's parent is a Body node, which can accept block-level nodes like Table.
        targetParagraph.ParentNode.InsertAfter(table, targetParagraph);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableInsertedAfterParagraph.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
