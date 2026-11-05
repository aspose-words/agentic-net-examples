using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the builder to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 50;
        textBox.WrapType = WrapType.Inline; // Ensure it stays within the header.
        textBox.StrokeColor = Color.Black;
        textBox.FillColor = Color.LightYellow;

        // Add a paragraph to the text box to hold text.
        Paragraph para = new Paragraph(doc);
        textBox.AppendChild(para);

        // Insert the text box into the header.
        builder.InsertNode(textBox);

        // Write text inside the text box.
        DocumentBuilder shapeBuilder = new DocumentBuilder(doc);
        shapeBuilder.MoveTo(textBox.FirstParagraph);
        shapeBuilder.Font.Size = 12;
        shapeBuilder.Writeln("First page header text box");

        // Save the document.
        string outputPath = "FirstPageHeaderWithTextBox.docx";
        doc.Save(outputPath);
    }
}
