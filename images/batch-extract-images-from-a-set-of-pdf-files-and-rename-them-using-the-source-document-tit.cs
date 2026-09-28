using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create deterministic sample images
        string[] sampleImagePaths = CreateSampleImages(baseDir);

        // Create sample PDF documents that contain the images and have titles
        CreateSamplePdfDocuments(inputFolder, sampleImagePaths);

        // Batch process PDFs: extract images and rename them using the document title
        int totalExtracted = 0;
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load PDF as Aspose.Words Document
            Document doc = new Document(pdfPath);

            // Determine source document title (fallback to file name without extension)
            string title = doc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            // Extract images from the document
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{title}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    if (!File.Exists(imagePath))
                        throw new Exception($"Failed to save extracted image: {imagePath}");
                    imageIndex++;
                    totalExtracted++;
                }
            }

            if (imageIndex == 1) // No images found in this document
                throw new Exception($"No images were extracted from PDF: {pdfPath}");
        }

        // Validate that at least one image was extracted overall
        if (totalExtracted == 0)
            throw new Exception("No images were extracted from any PDF files.");

        Console.WriteLine($"Extraction complete. Total images extracted: {totalExtracted}");
    }

    // Creates deterministic sample PNG images and returns their file paths
    private static string[] CreateSampleImages(string baseDir)
    {
        string[] paths = new string[2];
        for (int i = 0; i < 2; i++)
        {
            string filePath = Path.Combine(baseDir, $"sample{i + 1}.png");
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle with a distinct color
            using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.FromArgb(255, 0, 0, (i + 1) * 100), 5))
            {
                g.DrawRectangle(pen, 20, 20, 160, 160);
            }
            bitmap.Save(filePath);
            g.Dispose();
            bitmap.Dispose();
            paths[i] = filePath;
        }
        return paths;
    }

    // Generates sample PDF files containing the provided images and sets document titles
    private static void CreateSamplePdfDocuments(string inputFolder, string[] imagePaths)
    {
        for (int i = 0; i < imagePaths.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i + 1}");
            builder.InsertImage(imagePaths[i]);
            // Set the built‑in title property
            doc.BuiltInDocumentProperties.Title = $"SampleDoc{i + 1}";
            string pdfPath = Path.Combine(inputFolder, $"SampleDoc{i + 1}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}
