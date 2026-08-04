using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ApplyVignetteToJpegImages
{
    public static void Main()
    {
        // Directories for temporary files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string docsDir = Path.Combine(baseDir, "Docs");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample JPEG image that will be inserted into the documents.
        string sampleJpegPath = Path.Combine(baseDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 300, 200);

        // 2. Create a few Word documents and insert the sample JPEG image.
        int documentCount = 2;
        List<string> docPaths = new List<string>();
        for (int i = 0; i < documentCount; i++)
        {
            string docPath = Path.Combine(docsDir, $"Document_{i + 1}.docx");
            CreateDocumentWithJpeg(docPath, sampleJpegPath);
            docPaths.Add(docPath);
        }

        // 3. Process each document: extract JPEG images, apply vignette, replace them, and save results.
        int totalProcessedImages = 0;
        for (int docIndex = 0; docIndex < docPaths.Count; docIndex++)
        {
            string docPath = docPaths[docIndex];
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;
                if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

                // Extract the JPEG image to a memory stream.
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Apply vignette effect.
                    using (Bitmap originalBitmap = new Bitmap(originalStream))
                    {
                        ApplyVignette(originalBitmap);

                        // Save the modified image to a deterministic file name.
                        string vignetteFileName = Path.Combine(
                            outputDir,
                            $"vignette_doc{docIndex + 1}_img{imageIndex + 1}.jpg");
                        originalBitmap.Save(vignetteFileName, ImageFormat.Jpeg);

                        // Replace the image inside the shape with the modified one.
                        using (MemoryStream modifiedStream = new MemoryStream())
                        {
                            originalBitmap.Save(modifiedStream, ImageFormat.Jpeg);
                            modifiedStream.Position = 0;
                            shape.ImageData.SetImage(modifiedStream);
                        }

                        totalProcessedImages++;
                    }
                }

                imageIndex++;
            }

            // Save the modified document.
            string modifiedDocPath = Path.Combine(docsDir, $"Document_{docIndex + 1}_Modified.docx");
            doc.Save(modifiedDocPath);
        }

        // Validation: ensure at least one image was processed.
        if (totalProcessedImages == 0)
            throw new InvalidOperationException("No JPEG images were found and processed.");

        // The program finishes automatically; no user interaction required.
    }

    // Creates a simple JPEG image with a solid background and a colored rectangle.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.LightBlue);
            using (Brush brush = new SolidBrush(Color.Orange))
            {
                g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
            }
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a new Word document and inserts the specified JPEG image.
    private static void CreateDocumentWithJpeg(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document containing image: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple vignette effect by drawing concentric semi‑transparent black ellipses.
    private static void ApplyVignette(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Number of concentric ellipses.
            int steps = 10;
            // Maximum opacity (0‑255). 150 gives a noticeable darkening.
            int maxAlpha = 150;

            for (int i = 0; i < steps; i++)
            {
                float progress = (float)i / steps;
                int alpha = (int)(maxAlpha * progress);
                Color vignetteColor = Color.FromArgb(alpha, 0, 0, 0);
                using (Brush brush = new SolidBrush(vignetteColor))
                {
                    // Shrink the ellipse size on each step.
                    float insetX = width * progress * 0.5f;
                    float insetY = height * progress * 0.5f;
                    RectangleF rect = new RectangleF(insetX, insetY,
                        width - 2 * insetX, height - 2 * insetY);
                    g.FillEllipse(brush, rect);
                }
            }
        }
    }
}
