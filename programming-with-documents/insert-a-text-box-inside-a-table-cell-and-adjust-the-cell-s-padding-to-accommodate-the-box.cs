using System;
using System.IO;
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

        // Start a table and insert the first cell.
        builder.StartTable();
        Cell cell = builder.InsertCell();

        // Insert a text box shape into the current cell.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 50);

        // Move the builder's cursor inside the text box and add text.
        builder.MoveTo(textBox.LastParagraph);
        builder.Write("Hello from the text box!");

        // Adjust the cell padding so the text box is not clipped.
        cell.CellFormat.SetPaddings(10, 10, 10, 10);

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxInTableCell.docx");
        doc.Save(outputPath);
    }
}
