using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExtractBackgroundImages
{
    public static void Main()
    {
        // Prepare deterministic file names.
        string workingDir = Directory.GetCurrentDirectory();
        string backgroundImagePath = Path.Combine(workingDir, "bg.png");
        string documentPath = Path.Combine(workingDir, "sample.docx");
        string extractedImagePath = Path.Combine(workingDir, "background_extracted.png");

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be used as the document background.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with a solid color (light gray) for deterministic content.
                g.Clear(Aspose.Drawing.Color.LightGray);
            }

            // Save the image to a local file.
            bitmap.Save(backgroundImagePath);
        }

        // Verify that the image file was created.
        if (!File.Exists(backgroundImagePath))
            throw new Exception("Failed to create the sample background image.");

        // -----------------------------------------------------------------
        // 2. Create a new Word document and assign the image as its background shape.
        // -----------------------------------------------------------------
        Document doc = new Document();
        // Create a rectangle shape that will serve as the background.
        Shape backgroundShape = new Shape(doc, ShapeType.Rectangle);
        // Load the previously created image into the shape.
        backgroundShape.ImageData.SetImage(backgroundImagePath);
        // Assign the shape as the document's background.
        doc.BackgroundShape = backgroundShape;
        // Save the document to disk.
        doc.Save(documentPath);

        // Verify that the document file was created.
        if (!File.Exists(documentPath))
            throw new Exception("Failed to create the sample DOCX document.");

        // -----------------------------------------------------------------
        // 3. Load the document and extract the background image, if present.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        Shape bgShape = loadedDoc.BackgroundShape;

        if (bgShape == null || !bgShape.HasImage)
            throw new Exception("The document does not contain a background image.");

        // Save the extracted image as PNG.
        // Use FileFormatUtil to ensure correct extension based on image type.
        string extension = FileFormatUtil.ImageTypeToExtension(bgShape.ImageData.ImageType);
        // Force PNG extension regardless of original type for the task requirement.
        string finalPath = Path.ChangeExtension(extractedImagePath, ".png");
        bgShape.ImageData.Save(finalPath);

        // Validate that the extracted image file exists.
        if (!File.Exists(finalPath))
            throw new Exception("Failed to extract the background image.");

        // The program finishes here without any interactive prompts.
    }
}
