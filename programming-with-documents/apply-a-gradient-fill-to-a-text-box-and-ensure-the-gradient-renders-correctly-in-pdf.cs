using System;
using System.IO;
using System.Drawing;                     // For Color
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace GradientTextBoxExample
{
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
            textBox.Height = 150;
            textBox.Left = 100;
            textBox.Top = 100;

            // Apply a one‑color horizontal gradient fill to the text box.
            // The Fill object automatically sets the appropriate gradient angle for linear gradients,
            // so we do not set GradientAngle manually (setting it on a TextBox shape throws an exception).
            textBox.Fill.OneColorGradient(
                Color.LightBlue,               // Gradient foreground color.
                GradientStyle.Horizontal,      // Gradient direction.
                GradientVariant.Variant2,      // Gradient variant.
                0.2);                          // Lightness factor (0.0‑1.0).

            // Add a paragraph with a run of text inside the text box.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "Gradient Filled Text Box");
            para.AppendChild(run);
            textBox.AppendChild(para);

            // Insert the text box into the document body.
            builder.InsertNode(textBox);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document as PDF with DML rendering to preserve the gradient.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DmlRenderingMode = DmlRenderingMode.DrawingML
            };

            string pdfPath = Path.Combine(outputDir, "GradientTextBox.pdf");
            doc.Save(pdfPath, pdfOptions);
        }
    }
}
