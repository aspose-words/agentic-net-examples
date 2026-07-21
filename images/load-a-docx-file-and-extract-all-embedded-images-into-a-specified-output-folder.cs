using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Working directory for all temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a deterministic sample image.
        string sampleImagePath = Path.Combine(workDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // 2. Create a DOCX that contains the sample image.
        string docPath = Path.Combine(workDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Folder where extracted images will be saved.
        string outputFolder = Path.Combine(workDir, "ExtractedImages");
        Directory.CreateDirectory(outputFolder);

        // 4. Load the document and extract all embedded images.
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(outputFolder, $"image_{imageIndex}{extension}");

                // Save the image to the output folder.
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a blue rectangle as deterministic content.
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    graphics.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap to the specified file.
            bitmap.Save(path);
        }
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
