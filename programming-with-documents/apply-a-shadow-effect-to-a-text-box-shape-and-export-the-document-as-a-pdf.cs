using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

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
            Width = 300,
            Height = 100,
            WrapType = WrapType.None
        };

        // Apply a preset shadow effect to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow1; // any non‑mixed shadow type enables the shadow

        // Add a paragraph with some text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "This text box has a shadow.");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the shape into the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Save the document as a PDF file.
        doc.Save("ShadowedTextBox.pdf", SaveFormat.Pdf);
    }
}
