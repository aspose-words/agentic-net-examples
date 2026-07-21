using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

namespace AsposeWordsImageBatchExtract
{
    public class Program
    {
        public static void Main()
        {
            // Base directory for all generated files.
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(baseDir);

            // Folder that will contain the sample PDF files.
            string pdfInputDir = Path.Combine(baseDir, "InputPdfs");
            Directory.CreateDirectory(pdfInputDir);

            // Folder where extracted images will be saved.
            string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
            Directory.CreateDirectory(imageOutputDir);

            // -----------------------------------------------------------------
            // 1. Create a deterministic sample image that will be inserted into PDFs.
            // -----------------------------------------------------------------
            string sampleImagePath = Path.Combine(baseDir, "sample.png");
            CreateSampleImage(sampleImagePath);

            // -----------------------------------------------------------------
            // 2. Generate a few PDF files that contain the sample image and have a title.
            // -----------------------------------------------------------------
            const int pdfCount = 2;
            for (int i = 1; i <= pdfCount; i++)
            {
                CreatePdfWithImage(pdfInputDir, sampleImagePath, i);
            }

            // -----------------------------------------------------------------
            // 3. Batch process all PDFs: extract each image and rename using the document title.
            // -----------------------------------------------------------------
            string[] pdfFiles = Directory.GetFiles(pdfInputDir, "*.pdf");
            foreach (string pdfFile in pdfFiles)
            {
                ExtractImagesFromPdf(pdfFile, imageOutputDir);
            }

            // Simple validation: ensure at least one image was extracted.
            int totalExtracted = Directory.GetFiles(imageOutputDir).Length;
            if (totalExtracted == 0)
                throw new InvalidOperationException("No images were extracted from the PDF files.");

            // Program finishes automatically.
        }

        // Creates a 100x100 PNG image with a solid background.
        private static void CreateSampleImage(string filePath)
        {
            const int width = 100;
            const int height = 100;

            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Additional deterministic drawing can be added here if needed.
                bitmap.Save(filePath);
            }
        }

        // Creates a PDF file that contains the sample image and sets a title property.
        private static void CreatePdfWithImage(string outputDir, string imagePath, int index)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some text to make the page non‑empty.
            builder.Writeln($"This is sample PDF number {index}.");

            // Insert the deterministic image.
            builder.InsertImage(imagePath);

            // Set the document title which will be used for naming extracted images.
            doc.BuiltInDocumentProperties.Title = $"SamplePDF{index}";

            string pdfPath = Path.Combine(outputDir, $"Sample{index}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Loads a PDF, extracts all images, and saves them using the document title.
        private static void ExtractImagesFromPdf(string pdfPath, string outputDir)
        {
            Document pdfDoc = new Document(pdfPath);

            // Use the title property; fall back to file name if title is missing.
            string docTitle = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(docTitle))
                docTitle = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all shape nodes that may contain images.
            NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFile = Path.Combine(outputDir, $"{docTitle}_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(outputFile);
                imageIndex++;
            }

            // Validation: ensure at least one image was extracted from this PDF.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images found in PDF: {pdfPath}");
        }
    }
}
