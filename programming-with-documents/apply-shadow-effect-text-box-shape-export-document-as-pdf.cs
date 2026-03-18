using System;
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

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox.WrapType = WrapType.None; // make it float.

        // Add text inside the text box.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Shadowed text box");

        // Apply a shadow effect to the shape.
        textBox.ShadowFormat.Type = ShadowType.Shadow5; // any preset shadow.
        textBox.ShadowFormat.Color = Color.Gray;
        textBox.ShadowFormat.Transparency = 0.2; // 20% transparent.

        // Set up PDF save options (default configuration).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save("TextBoxWithShadow.pdf", pdfOptions);
    }
}
