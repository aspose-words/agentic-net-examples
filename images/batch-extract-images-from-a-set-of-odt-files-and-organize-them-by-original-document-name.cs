using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace BatchImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Define folders for sample data and output.
            string baseDir = Directory.GetCurrentDirectory();
            string inputDir = Path.Combine(baseDir, "InputDocs");
            string outputDir = Path.Combine(baseDir, "ExtractedImages");
            string tempImagePath = Path.Combine(baseDir, "sample.png");

            // Ensure clean environment.
            if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
            if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);

            // Create a deterministic sample image.
            CreateSampleImage(tempImagePath);

            // Create a few ODT documents each containing the sample image.
            for (int i = 1; i <= 3; i++)
            {
                string docName = $"Document{i}.odt";
                string docPath = Path.Combine(inputDir, docName);
                CreateOdtWithImage(docPath, tempImagePath);
            }

            // Batch extract images from all ODT files.
            foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
            {
                // Load the document.
                Document doc = new Document(odtFile);

                // Prepare output subfolder named after the source document (without extension).
                string docBaseName = Path.GetFileNameWithoutExtension(odtFile);
                string docOutputFolder = Path.Combine(outputDir, docBaseName);
                Directory.CreateDirectory(docOutputFolder);

                // Get all shape nodes that contain images.
                var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                    .Cast<Shape>()
                                    .Where(s => s.HasImage)
                                    .ToList();

                if (!shapeNodes.Any())
                    throw new InvalidOperationException($"No images found in document '{odtFile}'.");

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"Image_{imageIndex}{extension}";
                    string imagePath = Path.Combine(docOutputFolder, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imagePath);
                    imageIndex++;
                }
            }

            // Cleanup temporary sample image.
            if (File.Exists(tempImagePath))
                File.Delete(tempImagePath);
        }

        private static void CreateSampleImage(string filePath)
        {
            // 100x100 white background with a blue rectangle.
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Aspose.Drawing.Color.White);
                    using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Blue))
                    {
                        graphics.FillRectangle(brush, 10, 10, 80, 80);
                    }
                }
                bitmap.Save(filePath);
            }
        }

        private static void CreateOdtWithImage(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This document contains an image:");
            builder.InsertImage(imagePath);
            doc.Save(docPath, SaveFormat.Odt);
        }
    }
}
