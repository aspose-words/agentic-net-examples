using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the external image file.
        string imagePath = @"C:\Images\Sample.jpg";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Create a picture shape that links to the external image.
        // ------------------------------------------------------------
        Shape linkedPicture = new Shape(doc, ShapeType.Image);
        // By setting SourceFullName we make the picture a linked image.
        linkedPicture.ImageData.SourceFullName = imagePath;
        // Optional: set wrapping and size as needed.
        linkedPicture.WrapType = WrapType.Inline;

        // ------------------------------------------------------------
        // 2. Create a picture content control (StructuredDocumentTag).
        // ------------------------------------------------------------
        StructuredDocumentTag pictureControl = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Block);
        pictureControl.Title = "ExternalPictureControl";

        // Add the linked picture shape as a child of the content control.
        pictureControl.AppendChild(linkedPicture);

        // Insert the content control into the document at the current builder position.
        builder.InsertNode(pictureControl);

        // ------------------------------------------------------------
        // 3. Embed the linked image before saving.
        // ------------------------------------------------------------
        // Use Aspose.Words overload that accepts a file path – no System.Drawing needed.
        linkedPicture.ImageData.SetImage(imagePath);

        // ------------------------------------------------------------
        // 4. Save the document. The picture is now embedded.
        // ------------------------------------------------------------
        string outputPath = @"C:\Output\DocumentWithEmbeddedPicture.docx";
        doc.Save(outputPath);
    }
}
