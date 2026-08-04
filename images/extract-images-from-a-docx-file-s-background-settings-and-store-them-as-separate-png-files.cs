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
        const string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string backgroundImagePath = Path.Combine(artifactsDir, "bg.png");
        string documentPath = Path.Combine(artifactsDir, "BackgroundDoc.docx");
        string extractedImagePath = Path.Combine(artifactsDir, "extracted_background.png");

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with a solid color.
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle.
                using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }
            // Save the image to the file system.
            bitmap.Save(backgroundImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and set the background shape to the image.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a rectangle shape that will serve as the background.
        Shape backgroundShape = new Shape(doc, ShapeType.Rectangle);
        backgroundShape.ImageData.SetImage(backgroundImagePath);
        backgroundShape.Width = 500;   // Arbitrary size; actual size does not affect extraction.
        backgroundShape.Height = 500;

        // Assign the shape as the document's background.
        doc.BackgroundShape = backgroundShape;

        // Save the document.
        doc.Save(documentPath);

        // -----------------------------------------------------------------
        // 3. Load the document (optional) and extract the background image.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        Shape bgShape = loadedDoc.BackgroundShape;

        if (bgShape == null || !bgShape.HasImage)
            throw new InvalidOperationException("No background image found in the document.");

        // Determine appropriate file extension based on image type.
        string extension = FileFormatUtil.ImageTypeToExtension(bgShape.ImageData.ImageType);
        string outputPath = Path.ChangeExtension(extractedImagePath, extension);

        // Save the extracted image.
        bgShape.ImageData.Save(outputPath);

        // Validate that the image file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to extract the background image.", outputPath);
    }
}
