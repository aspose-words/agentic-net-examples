using System;
using System.IO;
using System.IO.Compression;
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
            // Root folder for all temporary files.
            string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtract");
            Directory.CreateDirectory(rootFolder);

            // Folder that will hold the sample DOCX files.
            string docsFolder = Path.Combine(rootFolder, "Docs");
            Directory.CreateDirectory(docsFolder);

            // Folder that will receive extracted images.
            string extractedFolder = Path.Combine(rootFolder, "ExtractedImages");
            Directory.CreateDirectory(extractedFolder);

            // Create deterministic sample images that will be inserted into the documents.
            string[] sampleImagePaths = CreateSampleImages(rootFolder);

            // Create a few sample DOCX files, each containing one of the sample images.
            CreateSampleDocuments(docsFolder, sampleImagePaths);

            // Process each DOCX file, extract its images and save them to the extracted folder.
            int totalExtracted = ExtractImagesFromDocuments(docsFolder, extractedFolder);

            // Validate that at least one image was extracted.
            if (totalExtracted == 0)
                throw new InvalidOperationException("No images were extracted from the documents.");

            // Create a zip archive that contains all extracted images.
            string zipPath = Path.Combine(rootFolder, "ExtractedImages.zip");
            if (File.Exists(zipPath))
                File.Delete(zipPath);
            ZipFile.CreateFromDirectory(extractedFolder, zipPath);

            // Validate that the zip file was created.
            if (!File.Exists(zipPath))
                throw new InvalidOperationException("Failed to create the zip archive.");

            // Example finished – all files are created on disk.
        }

        // Creates two deterministic PNG images using Aspose.Drawing and returns their file paths.
        private static string[] CreateSampleImages(string folder)
        {
            string[] paths = new string[2];

            for (int i = 0; i < 2; i++)
            {
                string filePath = Path.Combine(folder, $"sample{i + 1}.png");

                // Create a 100x100 bitmap.
                using (Bitmap bitmap = new Bitmap(100, 100))
                {
                    // Obtain a graphics object for drawing.
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        // Fill background with a distinct color.
                        Aspose.Drawing.Color fillColor = i == 0
                            ? Aspose.Drawing.Color.FromArgb(255, 200, 100)   // Light orange
                            : Aspose.Drawing.Color.FromArgb(100, 200, 255); // Light blue
                        graphics.Clear(fillColor);
                    }

                    // Save the bitmap to a PNG file.
                    bitmap.Save(filePath);
                }

                // Ensure the file exists.
                if (!File.Exists(filePath))
                    throw new InvalidOperationException($"Failed to create sample image: {filePath}");

                paths[i] = filePath;
            }

            return paths;
        }

        // Generates a couple of DOCX files, each containing one of the supplied images.
        private static void CreateSampleDocuments(string docsFolder, string[] imagePaths)
        {
            for (int i = 0; i < imagePaths.Length; i++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Insert the image into the document.
                builder.InsertImage(imagePaths[i]);

                string docPath = Path.Combine(docsFolder, $"Document{i + 1}.docx");
                doc.Save(docPath);

                // Validate that the document was saved.
                if (!File.Exists(docPath))
                    throw new InvalidOperationException($"Failed to save sample document: {docPath}");
            }
        }

        // Iterates over all DOCX files in the specified folder, extracts images, and returns the total count.
        private static int ExtractImagesFromDocuments(string docsFolder, string outputFolder)
        {
            int extractedCount = 0;
            string[] docFiles = Directory.GetFiles(docsFolder, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string docPath in docFiles)
            {
                // Load the document.
                Document doc = new Document(docPath);

                // Retrieve all shape nodes (including images).
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Determine the appropriate file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                    string imagePath = Path.Combine(outputFolder, imageFileName);

                    // Save the image to disk.
                    shape.ImageData.Save(imagePath);

                    // Validate that the image file was created.
                    if (!File.Exists(imagePath))
                        throw new InvalidOperationException($"Failed to save extracted image: {imagePath}");

                    extractedCount++;
                    imageIndex++;
                }
            }

            return extractedCount;
        }
    }
}
