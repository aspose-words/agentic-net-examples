using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert the image into the document. The method returns the Shape that represents the picture.
        Shape picture = builder.InsertImage(imagePath);

        // Example: set the picture to be an inline image (treated as a character in the text flow).
        picture.WrapType = WrapType.Inline;

        // Save the document in DOCX format.
        string outputPath = @"C:\Output\DocumentWithImage.docx";
        doc.Save(outputPath);
    }
}
