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
        // Base directory for temporary files.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Paths for the sample image, the document, and the folder for extracted images.
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        string sampleDocPath = Path.Combine(baseDir, "sample.docx");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure the output folder exists.
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100 white PNG).
        // -------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX file and embed the sample image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image from the file system.
        builder.InsertImage(sampleImagePath);
        // Save the document.
        doc.Save(sampleDocPath);

        // -------------------------------------------------
        // 3. Load the DOCX file and extract all embedded images.
        // -------------------------------------------------
        Document loadedDoc = new Document(sampleDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(extractedImagesDir, $"extracted_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Indicate success.
        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{extractedImagesDir}\".");
    }
}
