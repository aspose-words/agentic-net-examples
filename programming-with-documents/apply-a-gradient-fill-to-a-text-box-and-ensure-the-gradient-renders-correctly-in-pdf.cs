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
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 300;
        textBox.Height = 150;
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Center;

        // Apply a one‑color horizontal gradient fill to the text box.
        // Foreground color is LightBlue, gradient runs horizontally.
        textBox.Fill.OneColorGradient(Color.LightBlue, GradientStyle.Horizontal, GradientVariant.Variant2, 0.2);

        // Add a paragraph with centered text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Gradient Text Box");
        run.Font.Size = 24;
        run.Font.Bold = true;
        para.AppendChild(run);
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        textBox.AppendChild(para);

        // Insert the text box into the document body.
        builder.InsertNode(textBox);

        // Prepare output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as PDF. The gradient fill will be rendered in the PDF.
        string pdfPath = Path.Combine(outputDir, "GradientTextBox.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
