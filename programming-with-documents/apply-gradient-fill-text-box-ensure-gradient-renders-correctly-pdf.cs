using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 150);
        textBox.WrapType = WrapType.None; // keep it floating

        // Apply a two‑color horizontal gradient fill to the text box.
        textBox.Fill.TwoColorGradient(Color.LightBlue, Color.DarkBlue,
            GradientStyle.Horizontal, GradientVariant.Variant1);
        // GradientAngle works only for linear gradients; 0° = left‑to‑right.
        textBox.Fill.GradientAngle = 0;

        // Add some text inside the text box.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Size = 24;
        builder.Writeln("Gradient TextBox");

        // Save the document to PDF, rendering DrawingML shapes (including gradients) directly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML,
            DmlEffectsRenderingMode = DmlEffectsRenderingMode.Fine
        };
        doc.Save("GradientTextBox.pdf", pdfOptions);
    }
}
