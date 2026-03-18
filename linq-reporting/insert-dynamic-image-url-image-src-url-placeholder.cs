using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDynamicImage
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the image to be inserted.
        string imageUrl = "https://example.com/path/to/image.jpg";

        // Insert the image from the remote URL.
        // The InsertImage(string) overload automatically downloads the image.
        Shape imageShape = builder.InsertImage(imageUrl);

        // Optional: set hyperlink on the image (if desired).
        // imageShape.HRef = "https://example.com";

        // Save the document to a file.
        doc.Save("DynamicImage.docx");
    }
}
