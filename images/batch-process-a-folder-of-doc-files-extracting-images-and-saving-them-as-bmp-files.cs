using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "input.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(sampleImagePath);
        }

        // Create sample DOCX files that contain the sample image
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputFolder, $"SampleDoc{i}.docx");
            doc.Save(docPath);
        }

        // Batch process each DOC/DOCX file in the input folder
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string ext = Path.GetExtension(docFile).ToLowerInvariant();
            if (ext != ".doc" && ext != ".docx")
                continue;

            Document document = new Document(docFile);
            NodeCollection shapes = document.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                // Save the image data to a memory stream
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset stream before reading

                    // Load the image into Aspose.Drawing.Bitmap for BMP conversion
                    using (Bitmap bmp = new Bitmap(ms))
                    {
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}.bmp";
                        string outputPath = Path.Combine(outputFolder, outputFileName);
                        bmp.Save(outputPath, ImageFormat.Bmp);
                    }
                }

                imageIndex++;
            }
        }

        // Validate that at least one BMP file was created
        string[] bmpFiles = Directory.GetFiles(outputFolder, "*.bmp", SearchOption.TopDirectoryOnly);
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP images were extracted from the documents.");

        // Example completed without interactive prompts
    }
}
