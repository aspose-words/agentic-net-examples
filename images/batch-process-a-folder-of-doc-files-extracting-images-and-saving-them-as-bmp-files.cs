using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image to be used in the documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // Create a few sample DOCX files that contain the image
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"SampleDocument{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Batch process each DOC/DOCX file in the input folder
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                                            .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                        f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
        {
            Document doc = new Document(docFile);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            if (shapes.Count == 0)
                continue; // No images in this document

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Obtain the raw image bytes from the shape
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load the bytes into an Aspose.Drawing.Bitmap
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    ms.Position = 0; // Ensure stream is at the beginning
                    using (Bitmap bitmap = new Bitmap(ms))
                    {
                        // Build a deterministic BMP file name
                        string docName = Path.GetFileNameWithoutExtension(docFile);
                        string bmpFileName = $"{docName}_Image{imageIndex}.bmp";
                        string bmpPath = Path.Combine(outputFolder, bmpFileName);

                        // Save as BMP
                        bitmap.Save(bmpPath, ImageFormat.Bmp);
                    }
                }

                imageIndex++;
            }
        }

        // Simple validation: ensure at least one BMP was created
        int totalBmpFiles = Directory.GetFiles(outputFolder, "*.bmp", SearchOption.TopDirectoryOnly).Length;
        if (totalBmpFiles == 0)
            throw new InvalidOperationException("No images were extracted and saved as BMP files.");
    }

    // Creates a simple bitmap with a solid background and saves it to the specified path
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a DOCX document that inserts the provided image file
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times to have multiple shapes
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }
}
