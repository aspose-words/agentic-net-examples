using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders and file names.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string backgroundImagePath = Path.Combine(artifactsDir, "background.png");
        string documentPath = Path.Combine(artifactsDir, "docWithBackground.docx");
        string extractedImagePath = Path.Combine(artifactsDir, "extractedBackground.png");

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the image with a solid color.
                graphics.Clear(Color.LightBlue);
            }
            // Save the image to the file system.
            bitmap.Save(backgroundImagePath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and set the background shape.
        // -------------------------------------------------
        Document doc = new Document();

        // Create a rectangle shape that will serve as the background.
        Shape backgroundShape = new Shape(doc, ShapeType.Rectangle);
        // Assign the previously created image to the shape.
        backgroundShape.ImageData.SetImage(backgroundImagePath);
        // Optionally, set the shape size to match the page.
        backgroundShape.Width = doc.FirstSection.PageSetup.PageWidth;
        backgroundShape.Height = doc.FirstSection.PageSetup.PageHeight;

        // Apply the shape as the document's background.
        doc.BackgroundShape = backgroundShape;

        // Save the document.
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract the background image.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        Shape bgShape = loadedDoc.BackgroundShape;

        if (bgShape != null && bgShape.HasImage)
        {
            // Determine the appropriate file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(bgShape.ImageData.ImageType);
            // Ensure we save as PNG regardless of original type.
            string outputPath = Path.ChangeExtension(extractedImagePath, extension);
            bgShape.ImageData.Save(outputPath);

            // Validate that the image file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Failed to extract the background image.");
        }
        else
        {
            throw new InvalidOperationException("The document does not contain a background image.");
        }
    }
}
