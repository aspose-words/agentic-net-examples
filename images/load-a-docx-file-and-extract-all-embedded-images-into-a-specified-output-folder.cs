using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class Program
{
    public static void Main()
    {
        // Paths for the sample document and the folder where images will be extracted.
        string inputDocPath = "Sample.docx";
        string outputFolder = "ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // ------------------------------------------------------------
        // Step 1: Create a deterministic sample image (sample.png).
        // ------------------------------------------------------------
        string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 100;

        // Create a bitmap, clear it with white, and save it.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Optionally draw something simple (a black rectangle) to make the image non‑blank.
            graphics.DrawRectangle(new Pen(Color.Black, 2), 10, 10, imgWidth - 20, imgHeight - 20);
            bitmap.Save(sampleImagePath);
        }

        // ------------------------------------------------------------
        // Step 2: Build a DOCX that contains the sample image.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // ------------------------------------------------------------
        // Step 3: Load the DOCX and extract all embedded images.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(outputFolder, $"Image_{imageIndex}{extension}");

                // Save the image to the output folder.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Program finishes automatically; no user interaction required.
    }
}
