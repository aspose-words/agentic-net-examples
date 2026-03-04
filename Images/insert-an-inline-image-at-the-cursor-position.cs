using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder positioned at the start of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline image at the current cursor position.
        // Replace the path with the actual image file location or URL.
        string imagePath = @"C:\Images\Sample.jpg";
        builder.InsertImage(imagePath);

        // Save the document to the desired location.
        // Replace the path with the desired output file location.
        string outputPath = @"C:\Output\Result.docx";
        doc.Save(outputPath);
    }
}
