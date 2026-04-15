using System;
using System.IO;
using System.Drawing;                     // Needed for Color
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
        textBox.WrapType = WrapType.None;

        // Apply a two‑color horizontal gradient fill (blue → green).
        // Use the overload that accepts explicit colors.
        textBox.Fill.TwoColorGradient(
            Color.Blue,
            Color.Green,
            GradientStyle.Horizontal,
            GradientVariant.Variant1);

        // Add centered text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Gradient Text Box");
        run.Font.Size = 24;
        run.Font.Color = Color.White;
        para.AppendChild(run);
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        textBox.AppendChild(para);

        // Save the document as DOCX with strict OOXML compliance to retain gradient data.
        OoxmlSaveOptions docSaveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Strict
        };
        string outputDir = Path.Combine("Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "GradientTextBox.docx");
        doc.Save(docPath, docSaveOptions);

        // Convert the document to PDF. The gradient will render correctly.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        string pdfPath = Path.Combine(outputDir, "GradientTextBox.pdf");
        doc.Save(pdfPath, pdfSaveOptions);
    }
}
