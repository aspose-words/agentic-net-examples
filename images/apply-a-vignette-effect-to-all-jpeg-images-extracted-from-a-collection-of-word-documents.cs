using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample JPEG image.
        string sampleImagePath = "sample1.jpg";
        CreateSampleJpeg(sampleImagePath, 400, 300);

        // Create a folder for sample Word documents.
        string docsFolder = "Docs";
        Directory.CreateDirectory(docsFolder);

        // Create two Word documents each containing the sample JPEG image.
        string docPath1 = Path.Combine(docsFolder, "doc1.docx");
        CreateWordDocWithImage(docPath1, sampleImagePath);
        string docPath2 = Path.Combine(docsFolder, "doc2.docx");
        CreateWordDocWithImage(docPath2, sampleImagePath);

        // Folder for vignette‑processed images.
        string outputFolder = "VignetteOutput";
        Directory.CreateDirectory(outputFolder);

        int processedCount = 0;

        // Iterate over all DOCX files in the collection.
        foreach (string docFile in Directory.GetFiles(docsFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                // Process only JPEG images.
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
                {
                    // Save the extracted JPEG to a temporary file.
                    string tempImagePath = Path.Combine(outputFolder,
                        $"temp-{Path.GetFileNameWithoutExtension(docFile)}-{imageIndex}.jpg");
                    shape.ImageData.Save(tempImagePath);

                    // Apply vignette effect and save the result.
                    string vignettePath = Path.Combine(outputFolder,
                        $"vignette-{Path.GetFileNameWithoutExtension(docFile)}-{imageIndex}.jpg");
                    ApplyVignetteEffect(tempImagePath, vignettePath);

                    if (!File.Exists(vignettePath))
                        throw new Exception($"Failed to create vignette image: {vignettePath}");

                    // Clean up the temporary file.
                    File.Delete(tempImagePath);
                    processedCount++;
                    imageIndex++;
                }
            }
        }

        if (processedCount == 0)
            throw new Exception("No JPEG images were processed.");

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine($"Processed {processedCount} JPEG image(s) with vignette effect.");
    }

    // Generates a simple JPEG image with deterministic content.
    private static void CreateSampleJpeg(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document containing the specified image.
    private static void CreateWordDocWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple vignette effect by drawing concentric semi‑transparent ellipses.
    private static void ApplyVignetteEffect(string inputPath, string outputPath)
    {
        using (Bitmap original = new Bitmap(inputPath))
        {
            int width = original.Width;
            int height = original.Height;

            using (Bitmap result = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(result))
                {
                    // Draw the original image.
                    g.DrawImage(original, 0, 0, width, height);

                    // Overlay concentric ellipses with increasing opacity.
                    int steps = 10;
                    for (int i = 0; i < steps; i++)
                    {
                        float opacity = (float)(i + 1) / steps * 0.5f; // Max 50% opacity.
                        int inset = (int)(i * Math.Min(width, height) / (2.0 * steps));
                        Rectangle rect = new Rectangle(inset, inset, width - 2 * inset, height - 2 * inset);
                        using (SolidBrush brush = new SolidBrush(Color.FromArgb((int)(opacity * 255), 0, 0, 0)))
                        {
                            g.FillEllipse(brush, rect);
                        }
                    }
                }
                result.Save(outputPath, ImageFormat.Jpeg);
            }
        }
    }
}
