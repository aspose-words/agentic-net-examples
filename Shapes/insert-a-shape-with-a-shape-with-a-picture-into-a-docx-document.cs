using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document has a paragraph to host the shape.
        builder.Writeln();

        // Create a shape whose type is Image.
        Shape pictureShape = new Shape(builder.Document, ShapeType.Image);

        // Load an image from the local file system into the shape.
        // Replace the path with the actual location of your image file.
        pictureShape.ImageData.SetImage(@"C:\Images\SamplePicture.jpg");

        // Set the desired size of the shape (in points).
        pictureShape.Width = 150;   // 150 points wide
        pictureShape.Height = 150;  // 150 points high

        // Set the shape to be inline with the surrounding text.
        pictureShape.WrapType = WrapType.Inline;

        // Append the shape to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(pictureShape);

        // Save the document as a DOCX file.
        // Replace the path with the desired output location.
        doc.Save(@"C:\Output\ShapeWithPicture.docx");
    }
}
