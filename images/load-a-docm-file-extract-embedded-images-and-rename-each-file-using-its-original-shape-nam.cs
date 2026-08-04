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
        // Prepare directories.
        string baseDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(baseDir, "Data");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PNG images using Aspose.Drawing.
        string[] sampleImagePaths = new string[2];
        for (int i = 0; i < 2; i++)
        {
            string imagePath = Path.Combine(dataDir, $"sample{i + 1}.png");
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill with a solid color (different for each image).
                    Aspose.Drawing.Color fillColor = i == 0 ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightGreen;
                    g.Clear(fillColor);
                }
                bitmap.Save(imagePath);
            }
            sampleImagePaths[i] = imagePath;
        }

        // Create a DOCM document and insert the sample images.
        string docPath = Path.Combine(dataDir, "Sample.docm");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < sampleImagePaths.Length; i++)
        {
            // Insert image and obtain the Shape that represents it.
            Shape shape = builder.InsertImage(sampleImagePaths[i]);
            // Assign a deterministic name to the shape (used later for file naming).
            shape.Name = $"Image{i + 1}";
        }

        // Save the document as a macro-enabled DOCM file.
        doc.Save(docPath, SaveFormat.Docm);

        // Load the DOCM file.
        Document loadedDoc = new Document(docPath);

        // Extract all embedded images, renaming each file using its original shape name.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a file name based on the shape's name; fall back to an index if missing.
            string baseFileName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"Shape{extractedCount + 1}";
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outputPath = Path.Combine(outputDir, $"{baseFileName}{extension}");

            // Save the image data to the file system.
            shape.ImageData.Save(outputPath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{outputDir}'.");
    }
}
