using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing; // Aspose.Drawing for bitmap handling

public class Program
{
    public static void Main()
    {
        // Base directory of the executable.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Input folder that will contain sample DOC/DOCX files.
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Output folder where extracted BMP images will be saved.
        string outputFolder = Path.Combine(baseDir, "ExtractedBmp");
        Directory.CreateDirectory(outputFolder);

        // Create sample documents with images.
        CreateSampleDocuments(inputFolder);

        // Batch process all DOC/DOCX files in the input folder.
        ProcessDocuments(inputFolder, outputFolder);
    }

    // Creates a few DOCX files each containing a simple image.
    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 2; i++)
        {
            // Create a deterministic sample image using Aspose.Drawing.
            string imagePath = Path.Combine(folderPath, $"SampleImage{i}.png");
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid color.
                g.Clear(Color.FromArgb(50 * i, 100, 150));
                // Save the image to a file.
                bitmap.Save(imagePath);
            }

            // Create a new Word document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(imagePath);

            // Save the document.
            string docPath = Path.Combine(folderPath, $"Document{i}.docx");
            doc.Save(docPath);
        }
    }

    // Loads each document, extracts all images, and saves them as BMP files.
    private static void ProcessDocuments(string inputFolder, string outputFolder)
    {
        string[] docFiles = Directory.GetFiles(inputFolder, "*.doc*");

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);

            // Get all shape nodes that contain images.
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Save the image data to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading.

                    // Load the image into an Aspose.Drawing bitmap.
                    using (Bitmap bitmap = new Bitmap(imageStream))
                    {
                        // Construct a deterministic BMP file name.
                        string bmpFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.bmp";
                        string bmpPath = Path.Combine(outputFolder, bmpFileName);

                        // Save the bitmap as BMP.
                        bitmap.Save(bmpPath);
                    }
                }

                imageIndex++;
            }

            // Validation: ensure at least one image was extracted.
            if (imageIndex == 0)
            {
                throw new InvalidOperationException($"No images were extracted from '{docFile}'.");
            }
        }
    }
}
