using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace BatchImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Base folder for all sample data.
            string baseFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleData");
            string inputDocsFolder = Path.Combine(baseFolder, "InputDocs");
            string outputImagesFolder = Path.Combine(baseFolder, "ExtractedImages");

            // Ensure clean environment.
            if (Directory.Exists(baseFolder))
                Directory.Delete(baseFolder, true);
            Directory.CreateDirectory(inputDocsFolder);
            Directory.CreateDirectory(outputImagesFolder);

            // Create sample images of different formats.
            CreateSampleImage(Path.Combine(baseFolder, "sample_png.png"), 200, 200, Aspose.Drawing.Color.Red, Aspose.Drawing.Imaging.ImageFormat.Png);
            CreateSampleImage(Path.Combine(baseFolder, "sample_jpeg.jpg"), 200, 200, Aspose.Drawing.Color.Green, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
            CreateSampleImage(Path.Combine(baseFolder, "sample_bmp.bmp"), 200, 200, Aspose.Drawing.Color.Blue, Aspose.Drawing.Imaging.ImageFormat.Bmp);

            // Create sample documents each containing one of the images.
            CreateDocumentWithImage(Path.Combine(inputDocsFolder, "DocWithPng.docx"), Path.Combine(baseFolder, "sample_png.png"));
            CreateDocumentWithImage(Path.Combine(inputDocsFolder, "DocWithJpeg.docx"), Path.Combine(baseFolder, "sample_jpeg.jpg"));
            CreateDocumentWithImage(Path.Combine(inputDocsFolder, "DocWithBmp.docx"), Path.Combine(baseFolder, "sample_bmp.bmp"));

            // Process each document in the input folder.
            foreach (string docPath in Directory.GetFiles(inputDocsFolder, "*.docx"))
            {
                Document doc = new Document(docPath);
                NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;

                foreach (Shape shape in shapes)
                {
                    if (!shape.HasImage)
                        continue;

                    // Determine image format folder (png, jpeg, bmp, etc.).
                    string formatFolderName = shape.ImageData.ImageType.ToString().ToLowerInvariant();
                    string formatFolderPath = Path.Combine(outputImagesFolder, formatFolderName);
                    Directory.CreateDirectory(formatFolderPath);

                    // Build deterministic output file name.
                    string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}.{formatFolderName}";
                    string outputPath = Path.Combine(formatFolderPath, outputFileName);

                    // Save the image.
                    shape.ImageData.Save(outputPath);
                    imageIndex++;
                }

                if (imageIndex == 0)
                    throw new InvalidOperationException($"No images were extracted from document '{docPath}'.");
            }

            // Validation: ensure at least one image was saved.
            bool anyImageSaved = Directory.GetFiles(outputImagesFolder, "*.*", SearchOption.AllDirectories).Length > 0;
            if (!anyImageSaved)
                throw new InvalidOperationException("Image extraction completed but no images were saved.");
        }

        private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor, Aspose.Drawing.Imaging.ImageFormat imageFormat)
        {
            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(fillColor);
                }
                bitmap.Save(filePath, imageFormat);
            }
        }

        private static void CreateDocumentWithImage(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(imagePath);
            doc.Save(docPath);
        }
    }
}
