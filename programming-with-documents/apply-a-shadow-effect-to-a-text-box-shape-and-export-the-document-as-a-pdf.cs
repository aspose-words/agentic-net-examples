using System;
using System.IO;
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

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            WrapType = WrapType.None,
            Width = 300,
            Height = 100,
            Left = 100,
            Top = 100
        };

        // Apply a preset shadow to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow1;
        textBox.ShadowFormat.Color = Color.Gray;

        // Add a paragraph with some text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Text box with shadow");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the shape into the document.
        builder.InsertNode(textBox);

        // Define the output PDF file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextBoxShadow.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
