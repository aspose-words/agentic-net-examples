using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the sample cover image and the DOCX file.
        string coverImagePath = Path.Combine(artifactsDir, "cover.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");

        // 1. Create a deterministic sample image that will act as audio cover art.
        CreateSampleCoverImage(coverImagePath);

        // 2. Create a DOCX document and insert the cover image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape imageShape = builder.InsertImage(coverImagePath);
        imageShape.Width = 200;   // set explicit size
        imageShape.Height = 200;
        doc.Save(docPath);

        // 3. Load the document and extract all images (cover art) as JPEG files.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Save the image data to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;

                    // Load the image with Aspose.Drawing and re‑save it as JPEG.
                    using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(ms))
                    {
                        string outFile = Path.Combine(artifactsDir, $"CoverArt_{extractedCount}.jpg");
                        img.Save(outFile, ImageFormat.Jpeg);
                        extractedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple 200×200 PNG image with deterministic content.
    private static void CreateSampleCoverImage(string filePath)
    {
        int width = 200;
        int height = 200;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);

                // Draw deterministic text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    graphics.DrawString(
                        "Cover",
                        font,
                        Aspose.Drawing.Brushes.Black,
                        new Aspose.Drawing.PointF(20, 80));
                }
            }

            bitmap.Save(filePath);
        }
    }
}
