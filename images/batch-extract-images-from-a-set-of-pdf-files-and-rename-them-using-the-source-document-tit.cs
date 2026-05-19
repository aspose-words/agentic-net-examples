using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and extracted images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image to be used in the PDFs.
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Generate a few sample PDF files with titles and the sample image.
        GenerateSamplePdfs(inputFolder, sampleImagePath, 3);

        // Batch process each PDF: extract images and rename them using the document title.
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document doc = new Document(pdfPath);

            // Retrieve the document title; fall back to the file name without extension if missing.
            string title = doc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all shape nodes that contain images.
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .OfType<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            if (shapeNodes.Count == 0)
                throw new InvalidOperationException($"No images found in document '{pdfPath}'.");

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{title}_Image{imageIndex}{extension}";
                string imagePath = Path.Combine(outputFolder, imageFileName);

                // Save the image to the output folder.
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }
        }

        // Optional: indicate completion.
        Console.WriteLine("Image extraction completed.");
    }

    // Creates a simple white bitmap with optional drawing (currently just a white background).
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Additional deterministic drawing can be added here if desired.
            bitmap.Save(filePath);
        }
    }

    // Generates a specified number of PDF files, each containing the sample image and a title.
    private static void GenerateSamplePdfs(string folderPath, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the document title property.
            doc.BuiltInDocumentProperties.Title = $"SampleDoc{i}";

            // Insert the sample image.
            builder.InsertImage(imagePath);

            // Save as PDF.
            string pdfFileName = Path.Combine(folderPath, $"Doc{i}.pdf");
            doc.Save(pdfFileName, SaveFormat.Pdf);
        }
    }
}
