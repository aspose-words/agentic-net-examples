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
        textBox.WrapType = WrapType.None;               // No text wrapping.
        textBox.Left = 100;                             // Position from the left margin.
        textBox.Top = 100;                              // Position from the top margin.

        // Apply a horizontal one‑color gradient fill (red → lighter red).
        // GradientVariant.Variant2 gives a smooth transition.
        textBox.Fill.OneColorGradient(Color.Red, GradientStyle.Horizontal, GradientVariant.Variant2, 0.2);

        // Add a paragraph with some text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Gradient TextBox");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Save the document as DOCX (optional, for verification).
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "GradientTextBox.docx");
        doc.Save(docxPath);

        // Save the document as PDF, enabling high‑quality rendering to preserve the gradient.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseHighQualityRendering = true,                     // Improves rendering of gradients.
            DmlRenderingMode = DmlRenderingMode.DrawingML      // Ensures DrawingML shapes are kept.
        };
        string pdfPath = Path.Combine(outputDir, "GradientTextBox.pdf");
        doc.Save(pdfPath, pdfOptions);
    }
}
