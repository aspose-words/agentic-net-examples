using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractBackgroundImages
{
    public static void Main()
    {
        // Define file names.
        const string backgroundImagePath = "background.png";
        const string documentPath = "DocumentWithBackground.docx";
        const string extractedImagePath = "ExtractedBackground.png";

        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG).
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill with a solid color.
                graphics.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }

            // Save the image to the local file system.
            bitmap.Save(backgroundImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and set the background shape.
        // -------------------------------------------------
        Document doc = new Document();
        // Create a rectangle shape that will serve as the background.
        Shape backgroundShape = new Shape(doc, ShapeType.Rectangle);
        // Assign the previously created image to the shape.
        backgroundShape.ImageData.SetImage(backgroundImagePath);
        // Optionally adjust size to match the page.
        backgroundShape.Width = doc.FirstSection.PageSetup.PageWidth;
        backgroundShape.Height = doc.FirstSection.PageSetup.PageHeight;
        // Set the shape as the document background.
        doc.BackgroundShape = backgroundShape;

        // Save the document containing the background image.
        doc.Save(documentPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load the document and extract the background image.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        Shape bgShape = loadedDoc.BackgroundShape;

        if (bgShape == null || !bgShape.HasImage)
            throw new InvalidOperationException("No background image found in the document.");

        // Determine the appropriate file extension based on the image type.
        string extension = FileFormatUtil.ImageTypeToExtension(bgShape.ImageData.ImageType);
        // Ensure we save as PNG regardless of original format.
        string outputPath = Path.ChangeExtension(extractedImagePath, extension);

        // Save the extracted image.
        bgShape.ImageData.Save(outputPath);

        // -------------------------------------------------
        // 4. Validate that the image file was created.
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to extract the background image.", outputPath);

        // The program finishes here without any interactive prompts.
    }
}
