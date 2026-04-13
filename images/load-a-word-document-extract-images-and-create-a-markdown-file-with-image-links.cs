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
        // Define deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample image file (sample.png) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // Step 2: Create a Word document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateWordDocumentWithImage(docPath, sampleImagePath);

        // -----------------------------------------------------------------
        // Step 3: Load the Word document, extract all images, and build Markdown.
        // -----------------------------------------------------------------
        string markdownPath = Path.Combine(outputDir, "output.md");
        ExtractImagesAndCreateMarkdown(docPath, outputDir, markdownPath);
    }

    // Creates a deterministic PNG image file.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any previous file is removed.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and draw deterministic content.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle for visual distinction.
        graphics.DrawRectangle(new Pen(Color.Black, 5), 10, 10, width - 20, height - 20);
        graphics.Dispose();

        // Save the bitmap to the specified path.
        bitmap.Save(filePath, ImageFormat.Png);
        bitmap.Dispose();

        // Validate that the image file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    // Creates a Word document containing the specified image.
    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure any previous document is removed.
        if (File.Exists(docPath))
            File.Delete(docPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image using the builder (inline shape).
        Shape shape = builder.InsertImage(imagePath);
        shape.WrapType = WrapType.Inline;

        // Save the document.
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document file exists.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save Word document at '{docPath}'.");
    }

    // Loads the document, extracts images, and writes a Markdown file with image links.
    private static void ExtractImagesAndCreateMarkdown(string docPath, string imagesDir, string markdownPath)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Collect all shape nodes that contain images.
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

        if (shapeNodes.Count == 0)
            throw new InvalidOperationException("No images were found in the document.");

        // Prepare Markdown content.
        using (StreamWriter mdWriter = new StreamWriter(markdownPath, false))
        {
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);

                // Validate that the image file was created.
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save extracted image '{imageFullPath}'.");

                // Write Markdown image link.
                mdWriter.WriteLine($"![Image{imageIndex}]({imageFileName})");
                imageIndex++;
            }
        }

        // Validate that the Markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException($"Failed to create Markdown file at '{markdownPath}'.");
    }
}
