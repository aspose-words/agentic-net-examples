using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables; // Needed for Table class

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 150);
        // Make the shape floating and position it on the page.
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 100;   // points from the left edge of the page
        textBox.Top = 100;    // points from the top edge of the page

        // Move the builder's cursor into the text box so that subsequent content is added there.
        builder.MoveTo(textBox.FirstParagraph);

        // Build a 2x2 table inside the floating text box.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTextboxTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
