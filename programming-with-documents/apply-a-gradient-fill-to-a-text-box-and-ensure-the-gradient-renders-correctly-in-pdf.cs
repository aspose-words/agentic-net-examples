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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
        textBox.WrapType = WrapType.None; // Ensure the shape is floating.

        // Add a centered paragraph with some text inside the text box.
        Paragraph para = new Paragraph(doc);
        textBox.AppendChild(para);
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Gradient Text Box");
        para.AppendChild(run);

        // Apply a one‑color horizontal gradient fill to the text box.
        textBox.Fill.OneColorGradient(
            Color.LightBlue,               // Gradient foreground color.
            GradientStyle.Horizontal,      // Gradient direction.
            GradientVariant.Variant2,      // Gradient variant.
            0.2);                          // Lightness factor (0.0‑1.0).

        // Optionally set the gradient angle (not needed for horizontal style).
        textBox.Fill.GradientAngle = 0;

        // Prepare PDF save options. Using default options is sufficient for gradient rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Define output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GradientTextBox.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);
    }
}
