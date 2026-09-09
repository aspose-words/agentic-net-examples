using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string pdfDir = Path.Combine(artifactsDir, "PdfFiles");
        string imageOutputDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(pdfDir);
        Directory.CreateDirectory(imageOutputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be inserted into PDFs.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Generate a few sample PDF files containing the image.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            CreateSamplePdf(pdfDir, sampleImagePath, i);
        }

        // -----------------------------------------------------------------
        // 3. Batch process all PDFs: extract images and rename them using the document title.
        // -----------------------------------------------------------------
        string[] pdfFiles = Directory.GetFiles(pdfDir, "*.pdf");
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Determine a base name for extracted images (document title or file name).
            string docTitle = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(docTitle))
                docTitle = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all shapes that contain images.
            NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
            var imageShapes = shapeNodes
                .OfType<Shape>()
                .Where(s => s.HasImage)
                .ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in PDF '{pdfPath}'.");

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFile = Path.Combine(
                    imageOutputDir,
                    $"{docTitle}_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(outputFile);
                imageIndex++;
            }
        }

        // Simple validation: ensure at least one image was written.
        int totalExtracted = Directory.GetFiles(imageOutputDir).Length;
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the PDF files.");

        // Program finishes without waiting for user input.
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background.
            graphics.Clear(Color.LightBlue);

            // Draw a simple rectangle.
            graphics.DrawRectangle(new Pen(Color.DarkBlue, 5), 10, 10, width - 20, height - 20);

            // Save the bitmap.
            bitmap.Save(filePath);
        }
    }

    // Creates a PDF file with a title property and inserts the sample image.
    private static void CreateSamplePdf(string pdfFolder, string imagePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a title for later renaming of extracted images.
        doc.BuiltInDocumentProperties.Title = $"SampleDocument{index}";

        // Add some content.
        builder.Writeln($"This is sample PDF number {index}.");
        builder.InsertImage(imagePath);

        // Save as PDF.
        string pdfPath = Path.Combine(pdfFolder, $"Sample{index}.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
