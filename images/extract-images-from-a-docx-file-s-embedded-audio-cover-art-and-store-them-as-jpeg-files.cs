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
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a deterministic sample image that will act as audio cover art.
        string coverImagePath = Path.Combine(workDir, "cover.png");
        CreateSampleCoverImage(coverImagePath, 200, 200);

        // 2. Build a DOCX document and insert the cover image.
        string docPath = Path.Combine(workDir, "SampleWithAudio.docx");
        CreateDocumentWithCoverImage(docPath, coverImagePath);

        // 3. Extract all images (cover art) from the document and save them as JPEG files.
        ExtractImagesAsJpeg(docPath, workDir);

        // Validation – ensure at least one JPEG was created.
        int jpegCount = Directory.GetFiles(workDir, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted from the document.");

        // The example finishes without requiring user interaction.
    }

    // Creates a simple solid‑color PNG image using Aspose.Drawing.
    private static void CreateSampleCoverImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.CornflowerBlue);
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Generates a DOCX file that contains the previously created image.
    private static void CreateDocumentWithCoverImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image inline – in a real scenario this would be the audio object's cover art.
        builder.InsertImage(imagePath);

        doc.Save(docPath, SaveFormat.Docx);
    }

    // Extracts every image from the document, converts it to JPEG, and writes it to the output folder.
    private static void ExtractImagesAsJpeg(string docPath, string outputFolder)
    {
        Document doc = new Document(docPath);

        // Get all shape nodes (they may contain images).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image data to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image with Aspose.Drawing and re‑save it as JPEG.
                using (Bitmap bitmap = new Bitmap(imgStream))
                {
                    string jpegPath = Path.Combine(outputFolder, $"CoverArt_{imageIndex}.jpg");
                    bitmap.Save(jpegPath, ImageFormat.Jpeg);
                }
            }

            imageIndex++;
        }
    }
}
