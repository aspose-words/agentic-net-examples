using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Adjust the cell padding so the text box fits comfortably.
        builder.CellFormat.LeftPadding = 10;
        builder.CellFormat.RightPadding = 10;
        builder.CellFormat.TopPadding = 10;
        builder.CellFormat.BottomPadding = 10;

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 150;   // Width of the text box.
        textBox.Height = 80;   // Height of the text box.
        textBox.WrapType = WrapType.Inline; // Keep it inline within the cell.

        // Optional: set internal margins of the text box.
        textBox.TextBox.InternalMarginTop = 5;
        textBox.TextBox.InternalMarginBottom = 5;
        textBox.TextBox.InternalMarginLeft = 5;
        textBox.TextBox.InternalMarginRight = 5;

        // Add some text to the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Text inside the box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the current cell.
        builder.InsertNode(textBox);

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
