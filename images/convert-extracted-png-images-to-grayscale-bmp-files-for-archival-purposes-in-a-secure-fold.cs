using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = "Artifacts";
        string secureDir = "SecureArchive";
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(secureDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic PNG image that will be inserted into a document.
        // -----------------------------------------------------------------
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with a solid color.
            g.Clear(Color.CornflowerBlue);
            // Draw a simple red ellipse.
            g.FillEllipse(new SolidBrush(Color.Red), 50, 50, 100, 100);
            bitmap.Save(pngPath);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains the PNG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract PNG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image into an Aspose.Drawing.Bitmap.
                using (Bitmap sourceBmp = new Bitmap(imageStream))
                {
                    // Create a new bitmap for the grayscale version.
                    using (Bitmap grayBmp = new Bitmap(sourceBmp.Width, sourceBmp.Height))
                    using (Graphics graphics = Graphics.FromImage(grayBmp))
                    {
                        // Convert each pixel to grayscale.
                        for (int y = 0; y < sourceBmp.Height; y++)
                        {
                            for (int x = 0; x < sourceBmp.Width; x++)
                            {
                                Color srcColor = sourceBmp.GetPixel(x, y);
                                int grayValue = (int)(0.3 * srcColor.R + 0.59 * srcColor.G + 0.11 * srcColor.B);
                                Color grayColor = Color.FromArgb(grayValue, grayValue, grayValue);
                                grayBmp.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save the grayscale bitmap as BMP in the secure folder.
                        string outputPath = Path.Combine(secureDir, $"extracted_{extractedCount}.bmp");
                        grayBmp.Save(outputPath);

                        // Validate that the file was created.
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create '{outputPath}'.");
                    }
                }
            }

            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image was processed.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were found and converted.");
    }
}
