using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Centered Text");
        builder.EndRow();
        builder.EndTable();

        // Access the created table and its first cell.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell cell = table.Rows[0].Cells[0];

        // Set vertical alignment to center.
        cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Center the paragraph(s) inside the cell horizontally.
        foreach (Paragraph para in cell.Paragraphs)
        {
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
