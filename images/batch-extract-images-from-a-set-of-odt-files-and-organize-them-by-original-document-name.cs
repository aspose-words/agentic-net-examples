using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define folders for input ODT files and extracted images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");

        // Clean and recreate folders to ensure a deterministic run.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create a sample image that will be inserted into the ODT files.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(inputFolder, "sample.png");
        int imgWidth = 100;
        int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid color (white).
                g.Clear(Aspose.Drawing.Color.White);
            }
            // Save the bitmap to a PNG file.
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // Step 2: Create a few ODT documents that contain the sample image.
        // -----------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph and the sample image.
            builder.Writeln($"Document {docIndex} - first paragraph.");
            builder.InsertImage(sampleImagePath);
            builder.Writeln($"Document {docIndex} - second paragraph.");
            builder.InsertImage(sampleImagePath);

            // Save as ODT.
            string odtPath = Path.Combine(inputFolder, $"Doc{docIndex}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }

        // -----------------------------------------------------------------
        // Step 3: Batch process all ODT files, extracting images.
        // -----------------------------------------------------------------
        string[] odtFiles = Directory.GetFiles(inputFolder, "*.odt");
        foreach (string odtFile in odtFiles)
        {
            // Load the ODT document.
            Document doc = new Document(odtFile);

            // Get all shape nodes (including images).
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // Prepare a subfolder named after the original document (without extension).
            string docName = Path.GetFileNameWithoutExtension(odtFile);
            string docOutputFolder = Path.Combine(outputFolder, docName);
            Directory.CreateDirectory(docOutputFolder);

            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{docName}_image{imageIndex}{extension}";
                    string imagePath = Path.Combine(docOutputFolder, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imagePath);
                    imageIndex++;
                }
            }

            // Validation: ensure at least one image was extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{odtFile}'.");
        }

        // The program finishes without waiting for user input.
    }
}
