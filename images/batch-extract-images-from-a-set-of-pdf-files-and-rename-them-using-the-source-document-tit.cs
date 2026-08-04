using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing; // For Bitmap, Graphics, Color, Pen

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and extracted images.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        string tempImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean directories.
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder)) Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Step 1: Create a deterministic sample image.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        var bitmap = new Bitmap(imgWidth, imgHeight);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle for visual distinction.
        using var pen = new Pen(Color.Blue, 5);
        graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        // Save the image to a file that will be inserted into PDFs.
        bitmap.Save(tempImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // Step 2: Generate sample PDF files containing the image.
        // -------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Set a unique title for each document.
            doc.BuiltInDocumentProperties.Title = $"SampleDoc{i}";

            // Insert the sample image.
            builder.InsertImage(tempImagePath);

            // Save as PDF.
            string pdfPath = Path.Combine(inputFolder, $"Document{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // -------------------------------------------------
        // Step 3: Batch extract images from each PDF and rename using the document title.
        // -------------------------------------------------
        var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (var pdfFile in pdfFiles)
        {
            // Load the PDF document.
            var loadOptions = new PdfLoadOptions(); // default options
            var pdfDoc = new Document(pdfFile, loadOptions);

            // Retrieve the document title; fallback to file name without extension if empty.
            string title = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
                title = Path.GetFileNameWithoutExtension(pdfFile);

            // Collect all shapes that contain images.
            var imageShapes = pdfDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in PDF '{pdfFile}'.");

            int imageIndex = 0;
            foreach (var shape in imageShapes)
            {
                // Determine appropriate file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFileName = $"{title}_{imageIndex}{extension}";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the extracted image.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // -------------------------------------------------
        // Validation: ensure at least one image was extracted.
        // -------------------------------------------------
        int totalExtracted = Directory.GetFiles(outputFolder).Length;
        if (totalExtracted == 0)
            throw new InvalidOperationException("Image extraction failed: no images were saved.");

        // Program completed successfully.
    }
}
