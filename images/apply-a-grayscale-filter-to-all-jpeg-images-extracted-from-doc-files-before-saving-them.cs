using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Create a deterministic JPEG image using Aspose.Drawing.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                using (var brush = new SolidBrush(Color.Orange))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            // Explicitly specify JPEG format to guarantee the image type.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Build a DOCX document that contains the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        builder.InsertParagraph(); // Add a second image to demonstrate multiple extraction.
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "input.docx");
        doc.Save(docPath);

        // Reload the document (optional, demonstrates loading from file).
        Document loadedDoc = new Document(docPath);

        // Extract all JPEG images, apply grayscale, and save them.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original image into a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // Convert to grayscale pixel by pixel.
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            Color src = bitmap.GetPixel(x, y);
                            int gray = (int)(src.R * 0.3 + src.G * 0.59 + src.B * 0.11);
                            Color grayColor = Color.FromArgb(src.A, gray, gray, gray);
                            bitmap.SetPixel(x, y, grayColor);
                        }
                    }

                    // Determine output file name with proper extension.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outFile = Path.Combine(artifactsDir,
                        $"ExtractedImage_{imageIndex}{extension}");

                    // Save the processed (grayscale) image.
                    bitmap.Save(outFile, ImageFormat.Jpeg);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was saved.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and saved.");
    }
}
