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
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Width = 300;
        textBox.Height = 100;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph with a run of text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc);
        run.Text = "Shadowed Text Box";
        para.AppendChild(run);

        // Apply a preset shadow effect to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow1; // choose a shadow preset
        textBox.ShadowFormat.Color = Color.Gray;       // optional: set shadow color

        // Insert the shape into the document.
        builder.InsertParagraph(); // ensure there is a paragraph to host the shape
        builder.CurrentParagraph.AppendChild(textBox);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "TextBoxWithShadow.pdf");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
