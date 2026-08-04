using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

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
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top
        };

        // Add a paragraph with a run of text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Text box with shadow effect");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Apply a shadow effect to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow1; // preset shadow
        // The preset automatically makes the shadow visible.

        // Insert the shape into the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "ShadowTextBox.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
