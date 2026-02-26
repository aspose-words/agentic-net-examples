using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image. The method returns a Shape object representing the image.
        Shape imageShape = builder.InsertImage("ImageDir/Logo.jpg");

        // Lock the aspect ratio of the inserted image.
        bool lockAspectRatio = true;
        imageShape.AspectRatioLocked = lockAspectRatio;

        // Save the document.
        doc.Save("ArtifactsDir/AspectRatioLocked.docx");
    }
}
