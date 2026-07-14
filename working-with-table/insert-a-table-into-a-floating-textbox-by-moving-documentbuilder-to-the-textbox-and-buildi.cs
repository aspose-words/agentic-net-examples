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

        // Insert a floating textbox shape.
        Shape textbox = new Shape(doc, ShapeType.TextBox);
        textbox.Width = 300;               // Width in points.
        textbox.Height = 200;              // Height in points.
        textbox.WrapType = WrapType.None;  // Make the shape floating.
        textbox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textbox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textbox.HorizontalAlignment = HorizontalAlignment.Right;
        textbox.VerticalAlignment = VerticalAlignment.Bottom;

        // Insert the shape into the document at the current cursor position.
        builder.InsertNode(textbox);

        // Ensure the textbox contains at least one paragraph.
        if (textbox.FirstParagraph == null)
        {
            textbox.AppendChild(new Paragraph(doc));
        }

        // Move the builder to the first paragraph inside the textbox.
        builder.MoveTo(textbox.FirstParagraph);

        // Build a 2x2 table inside the floating textbox.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTextboxTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
