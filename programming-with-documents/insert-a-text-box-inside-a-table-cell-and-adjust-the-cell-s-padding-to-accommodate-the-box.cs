using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Increase the cell padding so the text box does not touch the cell borders.
        // Left, Top, Right, Bottom padding values are in points.
        builder.CellFormat.SetPaddings(10, 10, 10, 10);

        // Insert a text box shape into the current cell.
        // Width and height are also specified in points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 50);

        // Add a paragraph with some text inside the text box.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello inside the text box!");
        paragraph.AppendChild(run);
        textBox.AppendChild(paragraph);

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TextBoxInTableCell.docx");
        doc.Save(outputPath);
    }
}
