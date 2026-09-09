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

        // Insert a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Width = 300;
        textBox.Height = 100;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph with some text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Shadowed Text Box");
        para.AppendChild(run);

        // Apply a shadow effect to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow1; // preset shadow
        textBox.ShadowFormat.Color = Color.Gray;       // shadow color
        textBox.ShadowFormat.Transparency = 0.3;       // optional transparency

        // Insert the shape into the document.
        builder.InsertNode(textBox);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "TextBoxWithShadow.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
