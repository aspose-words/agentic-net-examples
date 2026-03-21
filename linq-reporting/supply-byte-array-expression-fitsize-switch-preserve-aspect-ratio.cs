using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertImageFromByteArray
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // PNG byte array (1x1 pixel). This avoids the need for System.Drawing.
        byte[] imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

        // Insert the image as an inline shape. The InsertImage method returns the Shape object.
        Shape imageShape = builder.InsertImage(imageBytes);

        // Preserve the original aspect ratio by fitting the image data to the shape frame.
        imageShape.ImageData.FitImageToShape();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ImageFromByteArray.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
