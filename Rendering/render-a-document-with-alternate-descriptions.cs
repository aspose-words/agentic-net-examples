using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class RenderDocumentWithAlternateDescriptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image that has alternate (alt) text.");

        // Insert an image and set its alternate text (description).
        // The image file should exist at the specified path.
        string imagePath = @"C:\Images\sample.png";
        builder.InsertImage(imagePath);
        // Retrieve the inserted shape (the image) and set its alternative text.
        var shape = (Aspose.Words.Drawing.Shape)doc.GetChild(NodeType.Shape, 0, true);
        shape.AlternativeText = "This is the alternate description for the image.";

        // Configure layout options to render comments as annotations (optional).
        // This demonstrates alternate rendering modes; not required for alt text.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Create PDF save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Ensure DrawingML shapes are rendered (default is DrawingML, but set explicitly).
        saveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

        // Save the document to PDF. The alternate text will be embedded in the PDF
        // and can be viewed in PDF readers that support alt text for images.
        string outputPath = @"C:\Output\DocumentWithAltText.pdf";
        doc.Save(outputPath, saveOptions);
    }
}
