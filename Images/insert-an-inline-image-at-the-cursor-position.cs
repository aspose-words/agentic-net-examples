using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline image at the current cursor position.
        // Replace the path with the actual location of your image file.
        string imagePath = "ImageDir/Logo.jpg";
        builder.InsertImage(imagePath);

        // Save the document to a file.
        doc.Save("Output/InlineImage.docx");
    }
}
