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

        // Start a new table.
        builder.StartTable();

        // Insert a cell and add some text.
        builder.InsertCell();
        builder.Writeln("Centered Text");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table, row, and cell.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell cell = table.Rows[0].Cells[0];

        // Set vertical alignment to center using CellFormat.
        cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Set horizontal alignment to center by aligning the paragraph inside the cell.
        Paragraph para = cell.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
