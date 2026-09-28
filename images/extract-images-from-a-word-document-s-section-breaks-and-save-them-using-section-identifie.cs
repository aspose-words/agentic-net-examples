using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

namespace ImageExtractionBySection
{
    public class Program
    {
        public static void Main()
        {
            // Prepare deterministic file names and folders
            string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workFolder);
            string inputImagePath = Path.Combine(workFolder, "sample.png");
            string documentPath = Path.Combine(workFolder, "sample.docx");
            string outputFolder = Path.Combine(workFolder, "ExtractedImages");
            Directory.CreateDirectory(outputFolder);

            // Create a sample image using Aspose.Drawing
            CreateSampleImage(inputImagePath);

            // Build a Word document with two sections, each containing an image
            BuildDocumentWithSections(documentPath, inputImagePath);

            // Load the document and extract images per section
            ExtractImagesBySection(documentPath, outputFolder);

            // Validate that at least one image was extracted
            int extractedCount = Directory.GetFiles(outputFolder, "*.png").Length;
            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // Optionally, clean up (comment out if you want to inspect files)
            // Directory.Delete(workFolder, true);
        }

        private static void CreateSampleImage(string filePath)
        {
            const int width = 200;
            const int height = 100;
            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    // Draw a simple rectangle
                    graphics.DrawRectangle(new Aspose.Drawing.Pen(Color.Blue, 3), 10, 10, width - 20, height - 20);
                }
                bitmap.Save(filePath);
            }

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Failed to create the sample image.", filePath);
        }

        private static void BuildDocumentWithSections(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First section with an image
            builder.InsertImage(imagePath);
            // Insert a section break (new page)
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section with the same image (could be different)
            builder.InsertImage(imagePath);

            // Save the document
            doc.Save(docPath, SaveFormat.Docx);

            if (!File.Exists(docPath))
                throw new FileNotFoundException("Failed to save the sample document.", docPath);
        }

        private static void ExtractImagesBySection(string docPath, string outputFolder)
        {
            Document doc = new Document(docPath);
            int totalExtracted = 0;

            for (int secIndex = 0; secIndex < doc.Sections.Count; secIndex++)
            {
                Section section = doc.Sections[secIndex];
                NodeCollection shapeNodes = section.Body.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;

                foreach (Shape shape in shapeNodes)
                {
                    if (shape.HasImage)
                    {
                        imageIndex++;
                        string outFile = Path.Combine(outputFolder,
                            $"section-{secIndex + 1}-image-{imageIndex}.png");
                        shape.ImageData.Save(outFile);
                        if (!File.Exists(outFile))
                            throw new InvalidOperationException($"Failed to save extracted image to {outFile}.");
                        totalExtracted++;
                    }
                }
            }

            if (totalExtracted == 0)
                throw new InvalidOperationException("No images were found in any section.");
        }
    }
}
