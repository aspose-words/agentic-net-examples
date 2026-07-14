using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

namespace BatchImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Root folder for the example
            string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageProcessing");
            string inputFolder = Path.Combine(rootFolder, "InputDocs");
            string imageFolder = Path.Combine(rootFolder, "ExtractedImages");
            string reportFolder = Path.Combine(rootFolder, "Report");

            // Ensure clean environment
            if (Directory.Exists(rootFolder))
                Directory.Delete(rootFolder, true);
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(imageFolder);
            Directory.CreateDirectory(reportFolder);

            // Create deterministic sample images
            string sampleImage1 = Path.Combine(rootFolder, "sample1.png");
            string sampleImage2 = Path.Combine(rootFolder, "sample2.png");
            CreateSampleImage(sampleImage1, 200, 150, Aspose.Drawing.Color.LightBlue);
            CreateSampleImage(sampleImage2, 120, 180, Aspose.Drawing.Color.LightCoral);

            // Create a few DOCX files that contain the sample images
            for (int docIndex = 1; docIndex <= 3; docIndex++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Insert one or two images depending on the document index
                builder.Writeln($"Document {docIndex}");
                builder.InsertImage(sampleImage1);
                if (docIndex % 2 == 0) // even documents get a second image
                    builder.InsertImage(sampleImage2);

                string docPath = Path.Combine(inputFolder, $"Document{docIndex}.docx");
                doc.Save(docPath);
            }

            // List to hold CSV rows
            List<string> csvLines = new List<string>();
            csvLines.Add("Document,ImageFile,ImageType,WidthPoints,HeightPoints,HorizontalResolution,VerticalResolution");

            int totalExtractedImages = 0;

            // Process each DOCX file in the input folder
            foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                Document doc = new Document(docPath);
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Determine file extension based on image type
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                    string imagePath = Path.Combine(imageFolder, imageFileName);

                    // Save the image to disk
                    shape.ImageData.Save(imagePath);
                    imageIndex++;
                    totalExtractedImages++;

                    // Gather metadata
                    ImageSize size = shape.ImageData.ImageSize;
                    string line = $"{Path.GetFileName(docPath)}," +
                                  $"{imageFileName}," +
                                  $"{shape.ImageData.ImageType}," +
                                  $"{size.WidthPoints:F2}," +
                                  $"{size.HeightPoints:F2}," +
                                  $"{size.HorizontalResolution:F2}," +
                                  $"{size.VerticalResolution:F2}";
                    csvLines.Add(line);
                }
            }

            // Validate that at least one image was extracted
            if (totalExtractedImages == 0)
                throw new InvalidOperationException("No images were extracted from the documents.");

            // Write CSV report
            string csvPath = Path.Combine(reportFolder, "ImageMetadataReport.csv");
            File.WriteAllLines(csvPath, csvLines);
        }

        // Helper method to create a deterministic bitmap image
        private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
        {
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(backgroundColor);
            // Additional deterministic drawing can be added here if needed
            bitmap.Save(filePath);
            graphics.Dispose();
            bitmap.Dispose();
        }
    }
}
