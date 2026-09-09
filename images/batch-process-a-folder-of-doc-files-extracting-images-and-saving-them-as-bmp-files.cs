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
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with images
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"));

        int totalExtracted = 0;

        // Process each DOC/DOCX file in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                                            .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                        f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
        {
            Document doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                using (MemoryStream imgStream = new MemoryStream())
                {
                    // Save the image data to a memory stream
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0; // Reset before reading

                    // Load the image into Aspose.Drawing.Bitmap
                    using (Bitmap bitmap = new Bitmap(imgStream))
                    {
                        // Ensure the bitmap is in a format that can be saved as BMP
                        string outFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.bmp";
                        string outPath = Path.Combine(outputFolder, outFileName);
                        bitmap.Save(outPath);
                        totalExtracted++;
                        imageIndex++;
                    }
                }
            }
        }

        // Validation: at least one BMP image must have been extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Example completed – the program exits automatically.
    }

    private static void CreateSampleDocument(string docPath)
    {
        // Create a deterministic sample image
        string imagePath = Path.ChangeExtension(docPath, ".png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            bitmap.Save(imagePath);
        }

        // Build a document and insert the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
