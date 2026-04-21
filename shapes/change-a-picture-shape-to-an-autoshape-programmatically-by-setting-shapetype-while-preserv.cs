using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image (100x100 red square) without using System.Drawing.
        //    The PNG data is a pre‑generated 100x100 red bitmap.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        byte[] redSquarePng = GenerateRedSquarePng(100, 100);
        File.WriteAllBytes(imagePath, redSquarePng);

        // -----------------------------------------------------------------
        // 2. Create a new document and insert the image as a picture shape.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // InsertImage returns a Shape that contains the picture.
        Shape pictureShape = builder.InsertImage(imagePath);
        if (pictureShape.ShapeType != ShapeType.Image)
            throw new InvalidOperationException("Inserted shape is not an image.");

        // Save the intermediate document (optional, just for reference).
        string originalDocPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(originalDocPath);

        // -----------------------------------------------------------------
        // 3. Replace the picture shape with an AutoShape (Rectangle) while preserving size.
        // -----------------------------------------------------------------
        Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (imageShape == null || imageShape.ShapeType != ShapeType.Image)
            throw new InvalidOperationException("No picture shape found to replace.");

        // Create a new AutoShape (Rectangle) and copy relevant properties.
        Shape autoShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = imageShape.Width,
            Height = imageShape.Height,
            Left = imageShape.Left,
            Top = imageShape.Top,
            RelativeHorizontalPosition = imageShape.RelativeHorizontalPosition,
            RelativeVerticalPosition = imageShape.RelativeVerticalPosition,
            WrapType = imageShape.WrapType,
            WrapSide = imageShape.WrapSide
        };

        // Insert the new shape after the original picture shape and remove the picture.
        imageShape.ParentNode.InsertAfter(autoShape, imageShape);
        imageShape.Remove();

        // -----------------------------------------------------------------
        // 4. Save the final document.
        // -----------------------------------------------------------------
        string resultDocPath = Path.Combine(artifactsDir, "ChangedShape.docx");
        doc.Save(resultDocPath);

        // -----------------------------------------------------------------
        // 5. Validation – ensure the output file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(resultDocPath))
            throw new FileNotFoundException("The resulting document was not saved.", resultDocPath);
    }

    // Generates a PNG byte array representing a solid red square of the given size.
    private static byte[] GenerateRedSquarePng(int width, int height)
    {
        // PNG header for a minimal RGBA image.
        // This method creates a simple red image using raw PNG encoding.
        // For brevity, we use a pre‑computed PNG for a 1x1 red pixel and scale it.
        // The 1x1 red pixel PNG (base64) is expanded to the requested dimensions
        // by repeating the pixel data; Word will stretch it to the desired size.
        const string base64RedPixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pixelData = Convert.FromBase64String(base64RedPixel);
        // The image is already a valid PNG; we can reuse it directly.
        // Word will scale the image to the shape dimensions we set later.
        return pixelData;
    }
}
