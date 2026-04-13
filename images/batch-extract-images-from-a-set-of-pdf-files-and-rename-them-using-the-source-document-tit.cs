using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchPdfImageExtractor
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string inputPdfDir = Path.Combine(baseDir, "InputPdfs");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputPdfDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be used in PDFs.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Generate a few sample PDF files containing the image.
        // -----------------------------------------------------------------
        CreateSamplePdf(Path.Combine(inputPdfDir, "DocumentOne.pdf"), "FirstDocument", sampleImagePath, 2);
        CreateSamplePdf(Path.Combine(inputPdfDir, "DocumentTwo.pdf"), "SecondDocument", sampleImagePath, 3);

        // -----------------------------------------------------------------
        // 3. Batch extract images from all PDFs and rename them using the document title.
        // -----------------------------------------------------------------
        int totalExtracted = 0;
        foreach (string pdfPath in Directory.GetFiles(inputPdfDir, "*.pdf"))
        {
            // Load PDF with default options (images are not skipped).
            var loadOptions = new PdfLoadOptions();
            Document doc = new Document(pdfPath, loadOptions);

            // Determine a title for naming extracted images.
            string title = doc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all shape nodes that may contain images.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in PDF '{pdfPath}'.");

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Build a deterministic file name: {Title}_Image_{index}{extension}
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFileName = $"{title}_Image_{imageIndex}{extension}";
                string outputPath = Path.Combine(artifactsDir, outputFileName);

                // Save the image.
                shape.ImageData.Save(outputPath);
                imageIndex++;
                totalExtracted++;
            }
        }

        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from any PDF.");

        Console.WriteLine("Image extraction completed.");
    }

    // Creates a simple white image with optional drawing (deterministic).
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any previous file is removed.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and graphics using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Additional deterministic drawing can be added here if desired.
            bitmap.Save(filePath);
        }
    }

    // Creates a PDF document with a given title and inserts the specified image a number of times.
    private static void CreateSamplePdf(string pdfPath, string title, string imagePath, int repeatCount)
    {
        // Ensure any previous file is removed.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document title property.
        doc.BuiltInDocumentProperties.Title = title;

        for (int i = 0; i < repeatCount; i++)
        {
            // Insert the image using the builder (creates a Shape internally and inserts it).
            Shape shape = builder.InsertImage(imagePath);
            // No need to append the shape again; InsertImage already adds it to the document.

            // Add a line break between images.
            builder.Writeln();
        }

        // Save as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
