using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        const string sampleImagePath = "input.png";
        const string docmPath = "sample.docm";

        // -------------------------------------------------
        // Create a sample image using Aspose.Drawing
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // (Optional) draw something deterministic
        graphics.Dispose();
        bitmap.Save(sampleImagePath);
        bitmap.Dispose();

        // -------------------------------------------------
        // Create a DOCM document and insert the image as a shape
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape imageShape = builder.InsertImage(sampleImagePath);
        // Assign a deterministic shape name
        imageShape.Name = "SampleImageShape";
        // Save the document as DOCM
        doc.Save(docmPath, SaveFormat.Docm);

        // -------------------------------------------------
        // Load the DOCM document and extract embedded images
        // -------------------------------------------------
        Document loadedDoc = new Document(docmPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Determine a safe file name based on the shape's name
            string baseName = string.IsNullOrWhiteSpace(shape.Name)
                ? $"image_{extractedCount}"
                : shape.Name;

            // Remove invalid file name characters
            foreach (char c in Path.GetInvalidFileNameChars())
                baseName = baseName.Replace(c, '_');

            // Determine file extension from image type
            string extension = shape.ImageData.ImageType.ToString().ToLowerInvariant();
            string outputFileName = $"{baseName}.{extension}";

            // Save the image to disk
            shape.ImageData.Save(outputFileName);
            extractedCount++;
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional: indicate success
        Console.WriteLine($"Extracted {extractedCount} image(s) from '{docmPath}'.");
    }
}
