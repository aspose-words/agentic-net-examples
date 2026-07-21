using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class ExtractBackgroundImages
{
    public static void Main()
    {
        // Prepare folders and file names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string backgroundImagePath = Path.Combine(artifactsDir, "background.png");
        string documentPath = Path.Combine(artifactsDir, "DocumentWithBackground.docx");
        string extractedImagePath = Path.Combine(artifactsDir, "ExtractedBackground.png");

        // -------------------------------------------------
        // 1. Create a deterministic sample image (200x200).
        // -------------------------------------------------
        Bitmap bitmap = new Bitmap(200, 200);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Draw a simple rectangle to make the image recognizable.
        graphics.DrawRectangle(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
        graphics.Dispose();
        bitmap.Save(backgroundImagePath);
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a DOCX document and set its background.
        // -------------------------------------------------
        Document doc = new Document();

        // The background must be a rectangle shape.
        Shape backgroundShape = new Shape(doc, ShapeType.Rectangle);
        backgroundShape.ImageData.SetImage(backgroundImagePath);
        // Assign the shape as the document's background.
        doc.BackgroundShape = backgroundShape;

        // Save the document to the file system.
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract the background image.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        Shape bgShape = loadedDoc.BackgroundShape;

        if (bgShape != null && bgShape.HasImage)
        {
            // Save the extracted image as PNG.
            bgShape.ImageData.Save(extractedImagePath);
        }

        // -------------------------------------------------
        // 4. Validate that the image was written.
        // -------------------------------------------------
        if (!File.Exists(extractedImagePath))
            throw new InvalidOperationException("Failed to extract the background image.");

        // The program finishes here without any interactive prompts.
    }
}
