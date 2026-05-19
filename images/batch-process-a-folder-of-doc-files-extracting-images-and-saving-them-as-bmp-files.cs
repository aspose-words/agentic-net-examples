using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define folders for input documents and extracted images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");

        // Ensure clean folders.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        Directory.CreateDirectory(inputFolder);

        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image that will be inserted into the documents.
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // Create a few sample DOC files containing the image.
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputFolder, $"SampleDocument{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Batch process each DOC/DOCX file in the input folder.
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.doc*"))
        {
            // Load the document.
            Document doc = new Document(docFile);

            // Collect all shape nodes.
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Save the shape's image data to a memory stream.
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading.

                    // Load the image into an Aspose.Drawing.Bitmap.
                    using (Bitmap bitmap = new Bitmap(imgStream))
                    {
                        // Ensure the bitmap is in a format that can be saved as BMP.
                        // Save the bitmap as BMP to the output folder.
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}.bmp";
                        string outputPath = Path.Combine(outputFolder, outputFileName);
                        bitmap.Save(outputPath, ImageFormat.Bmp);
                        imageIndex++;
                    }
                }
            }

            // Validation: ensure at least one image was extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were found in document '{docFile}'.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Image extraction completed.");
    }

    // Creates a simple PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            graphics.DrawRectangle(new Pen(Aspose.Drawing.Color.DarkBlue, 3), 10, 10, width - 20, height - 20);
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX document and inserts the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
