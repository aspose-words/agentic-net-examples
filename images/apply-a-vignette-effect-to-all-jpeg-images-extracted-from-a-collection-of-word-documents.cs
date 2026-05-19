using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;      // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string tempImagesDir = Path.Combine(artifactsDir, "TempImages");
        Directory.CreateDirectory(tempImagesDir);

        // 1. Create a sample JPEG image that will be inserted into the documents.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 200, 200);

        // 2. Create a collection of Word documents that contain the JPEG image.
        int documentCount = 2;
        string[] docPaths = new string[documentCount];
        for (int i = 0; i < documentCount; i++)
        {
            string docPath = Path.Combine(artifactsDir, $"Document{i + 1}.docx");
            CreateDocumentWithImage(docPath, sampleJpegPath);
            docPaths[i] = docPath;
        }

        // 3. Process each document: extract JPEG images, apply vignette, replace them, and save the document.
        int totalVignetteImages = 0;
        for (int docIndex = 0; docIndex < docPaths.Length; docIndex++)
        {
            Document doc = new Document(docPaths[docIndex]);

            // Collect all shape nodes that contain JPEG images.
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                            .ToList();

            int imageIndex = 0;
            foreach (var shape in shapes)
            {
                // a) Save the original JPEG image to a temporary file.
                string originalImagePath = Path.Combine(tempImagesDir,
                    $"doc{docIndex + 1}_img{imageIndex}.jpg");
                shape.ImageData.Save(originalImagePath);

                // b) Apply vignette effect.
                string vignetteImagePath = Path.Combine(tempImagesDir,
                    $"doc{docIndex + 1}_img{imageIndex}_vignette.jpg");
                ApplyVignetteEffect(originalImagePath, vignetteImagePath);
                totalVignetteImages++;

                // c) Replace the image in the shape with the vignette version.
                shape.ImageData.SetImage(vignetteImagePath);

                imageIndex++;
            }

            // d) Save the modified document.
            string outputDocPath = Path.Combine(artifactsDir,
                $"Document{docIndex + 1}_Vignette.docx");
            doc.Save(outputDocPath);
        }

        // Validation: ensure at least one vignette image was produced.
        if (totalVignetteImages == 0)
            throw new InvalidOperationException("No JPEG images were found to apply the vignette effect.");

        // Cleanup temporary files (optional).
        // Directory.Delete(tempImagesDir, true);
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                g.Clear(Color.LightBlue);

                // Draw a simple rectangle.
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 100, 150, 200)))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }

            // Save as JPEG using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document that contains the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image.
        builder.InsertImage(imagePath);

        // Add a paragraph after the image for clarity.
        builder.Writeln();

        doc.Save(docPath);
    }

    // Applies a simple vignette effect to the source JPEG and writes the result to the destination path.
    private static void ApplyVignetteEffect(string sourcePath, string destinationPath)
    {
        using (Bitmap sourceBitmap = new Bitmap(sourcePath))
        {
            int width = sourceBitmap.Width;
            int height = sourceBitmap.Height;

            using (Bitmap resultBitmap = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(resultBitmap))
                {
                    // Draw the original image.
                    g.DrawImage(sourceBitmap, 0, 0, width, height);

                    // Parameters for the vignette.
                    int steps = 10;
                    int maxRadius = (int)Math.Sqrt(width * width + height * height) / 2;

                    // Draw concentric semi‑transparent black ellipses, darker at the edges.
                    for (int i = 0; i < steps; i++)
                    {
                        float t = (float)i / steps; // 0 = outermost, 1 = innermost
                        int radius = (int)(maxRadius * (1 - t));
                        int alpha = (int)(150 * t); // 0 at center, 150 at edges

                        using (SolidBrush brush = new SolidBrush(Color.FromArgb(alpha, 0, 0, 0)))
                        {
                            g.FillEllipse(brush,
                                width / 2 - radius,
                                height / 2 - radius,
                                radius * 2,
                                radius * 2);
                        }
                    }
                }

                // Save the result as JPEG using Aspose.Drawing.Imaging.ImageFormat.
                resultBitmap.Save(destinationPath, ImageFormat.Jpeg);
            }
        }
    }
}
