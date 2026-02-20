using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image and obtain the Shape object that represents it.
        Shape shape = builder.InsertImage("ImageDir/Logo.jpg");

        // Make the shape floating (not inline) and place it behind the text.
        shape.WrapType = WrapType.None;
        shape.BehindText = true;

        // Position the shape relative to the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Center the shape on the page.
        shape.HorizontalAlignment = HorizontalAlignment.Center;
        shape.VerticalAlignment = VerticalAlignment.Center;

        // Save the document as a DOCX file.
        doc.Save("ArtifactsDir/Image.InsertShape.docx");
    }
}
