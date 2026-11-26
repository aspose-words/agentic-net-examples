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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Centered Text");

        // Retrieve the cell that was just created.
        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

        // Set vertical alignment of the cell to center.
        cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Align all paragraphs inside the cell horizontally to center.
        foreach (Paragraph para in cell.Paragraphs)
        {
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        }

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "AlignedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
