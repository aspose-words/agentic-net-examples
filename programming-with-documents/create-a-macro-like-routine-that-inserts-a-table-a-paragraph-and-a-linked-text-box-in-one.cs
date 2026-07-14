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

        // ---------- Insert a table ----------
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // ---------- Insert a paragraph ----------
        builder.Writeln("This paragraph follows the table.");

        // ---------- Insert a linked text box ----------
        // Insert a free‑floating text box shape.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,
            RelativeHorizontalPosition.Margin, 0,
            RelativeVerticalPosition.Margin, 0,
            200, // width in points
            100, // height in points
            WrapType.None);

        // Add text to the text box.
        textBox.FirstParagraph.AppendChild(new Run(doc, "Content of the linked text box."));

        // Set the hyperlink for the shape using the HRef property.
        textBox.HRef = "https://www.example.com";
        // Optional: open the link in a new window and set a screen tip.
        textBox.Target = "New Window";
        textBox.ScreenTip = "Open example.com";

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroLikeOutput.docx");
        doc.Save(outputPath);
    }
}
