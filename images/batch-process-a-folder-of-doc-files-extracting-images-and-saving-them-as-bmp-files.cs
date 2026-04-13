using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder that will contain sample DOC files.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Folder where extracted BMP images will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create sample DOC files with images (required because we cannot
        // assume any external files exist).
        // -----------------------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            // Create a deterministic sample image (PNG) using Aspose.Drawing.
            string sampleImagePath = Path.Combine(docsFolder, $"sample{i}.png");
            CreateSampleImage(sampleImagePath, 200, 150, Aspose.Drawing.Color.LightBlue, $"Img {i}");

            // Build a document and insert the sample image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            builder.Writeln($"Document {i} with an image.");

            // Save as a legacy DOC file.
            string docPath = Path.Combine(docsFolder, $"Document{i}.doc");
            doc.Save(docPath, SaveFormat.Doc);
        }

        // -----------------------------------------------------------------
        // Batch process each DOC file: extract all images and save as BMP.
        // -----------------------------------------------------------------
        string[] docFiles = Directory.GetFiles(docsFolder, "*.doc");
        foreach (string docFile in docFiles)
        {
            Document document = new Document(docFile);
            NodeCollection shapeNodes = document.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Save the image data to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading.

                    // Load the image with Aspose.Drawing and re‑save as BMP.
                    using (Aspose.Drawing.Image asposeImage = Aspose.Drawing.Image.FromStream(imageStream))
                    {
                        string bmpFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.bmp";
                        string bmpPath = Path.Combine(outputFolder, bmpFileName);
                        asposeImage.Save(bmpPath);
                        imageIndex++;
                    }
                }
            }

            // Validation: at least one image must have been extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were found in document '{docFile}'.");
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine("Image extraction completed.");
    }

    // Creates a simple PNG image with a solid background and optional text.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                // Draw deterministic text if needed.
                // (Font usage is avoided as it's not required for the task.)
            }

            bitmap.Save(filePath);
        }
    }
}
