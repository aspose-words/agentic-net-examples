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

        // Insert an image into the document; the method returns a Shape object.
        Shape shape = builder.InsertImage("ImageDir\\Logo.jpg");

        // Optional: configure the shape's layout and alignment.
        shape.WrapType = WrapType.Square; // Text wraps around the shape.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page; // Position relative to the page.
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;   // Position relative to the page.
        shape.HorizontalAlignment = HorizontalAlignment.Center; // Center horizontally.
        shape.VerticalAlignment = VerticalAlignment.Center;     // Center vertically.

        // Save the document to a DOCX file.
        doc.Save("ArtifactsDir\\ImageShape.docx");
    }
}
