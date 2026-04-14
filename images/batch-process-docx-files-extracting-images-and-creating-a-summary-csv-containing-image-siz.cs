using System;
using System.IO;
using System.Text;
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
        // Base working directory.
        string baseDir = Directory.GetCurrentDirectory();

        // -----------------------------------------------------------------
        // Step 1: Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // Step 2: Create sample DOCX files that contain the sample image.
        // -----------------------------------------------------------------
        string docsDir = Path.Combine(baseDir, "Docs");
        Directory.CreateDirectory(docsDir);

        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the sample image twice into each document.
            builder.Writeln($"Document {docIndex}");
            builder.InsertImage(sampleImagePath);
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertImage(sampleImagePath);

            string docPath = Path.Combine(docsDir, $"Sample{docIndex}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // Step 3: Batch process all DOCX files, extract images, and build CSV.
        // -----------------------------------------------------------------
        string imagesOutDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(imagesOutDir);

        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Document,ImageFile,WidthPixels,HeightPixels,ImageType");

        int totalExtractedImages = 0;

        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                string imageOutPath = Path.Combine(imagesOutDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageOutPath);

                // Retrieve image dimensions in pixels.
                int widthPx = shape.ImageData.ImageSize.WidthPixels;
                int heightPx = shape.ImageData.ImageSize.HeightPixels;

                // Append a line to the CSV summary.
                csvBuilder.AppendLine($"{Path.GetFileName(docFile)},{imageFileName},{widthPx},{heightPx},{shape.ImageData.ImageType}");

                imageIndex++;
                totalExtractedImages++;
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Write the CSV summary file.
        string csvPath = Path.Combine(baseDir, "summary.csv");
        File.WriteAllText(csvPath, csvBuilder.ToString());

        // -----------------------------------------------------------------
        // End of processing. All files are written to disk.
        // -----------------------------------------------------------------
    }

    // -----------------------------------------------------------------
    // Helper: creates a deterministic PNG image using Aspose.Drawing.
    // -----------------------------------------------------------------
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any existing file is overwritten.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and draw a simple pattern.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple black rectangle.
        using (Pen pen = new Pen(Color.Black, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap as PNG.
        bitmap.Save(filePath, ImageFormat.Png);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }
}
