using System;
using System.IO;
using System.Drawing; // Required for Color
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class GradientTextBoxExample
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

        // Apply a one‑color vertical gradient fill to the text box.
        // Foreground color: LightBlue, style: Vertical, variant: Variant1, degree: 0.2 (light to dark).
        textBox.Fill.OneColorGradient(Color.LightBlue, GradientStyle.Vertical, GradientVariant.Variant1, 0.2);

        // Add a paragraph with a run of text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Gradient Filled Text Box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the document body.
        builder.InsertNode(textBox);

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GradientTextBox.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document as PDF. Use PdfSaveOptions to keep drawing fidelity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render DrawingML (DML) shapes to preserve gradient details.
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };
        doc.Save(outputPath, pdfOptions);
    }
}
