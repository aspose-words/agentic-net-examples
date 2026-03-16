using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Iterate through all shapes that may contain images.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Skip shapes without an image.
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Original image width in pixels.
                int originalWidthPx = shape.ImageData.ImageSize.WidthPixels;

                // If the width exceeds 800 pixels, resize proportionally.
                if (originalWidthPx > 800)
                {
                    double scaleFactor = 800.0 / originalWidthPx;

                    // Shape dimensions are in points (1 point = 1/72 inch).
                    // Scaling both width and height preserves the aspect ratio.
                    shape.Width  *= scaleFactor;
                    shape.Height *= scaleFactor;
                }
            }
        }

        // Save the document with resized images.
        doc.Save("output.docx");
    }
}
