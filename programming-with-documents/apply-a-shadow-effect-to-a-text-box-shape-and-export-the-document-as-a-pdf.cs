using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a blank document.
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

        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Shadowed Text Box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Apply a shadow effect to the shape.
        // Setting the Type automatically makes the shadow visible.
        textBox.ShadowFormat.Type = ShadowType.Shadow5; // preset shadow type
        textBox.ShadowFormat.Color = Color.Gray;
        textBox.ShadowFormat.Transparency = 0.3; // 30% transparent

        // Insert the shape into the document.
        builder.InsertNode(textBox);

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "TextBoxShadow.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
