using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample image (PNG) to be inserted into documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files containing the image
        CreateSampleDocuments(inputDir, sampleImagePath, 2);

        // Batch process: extract images from each DOCX and save as BMP
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string bmpFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.bmp";
                    string bmpPath = Path.Combine(outputDir, bmpFileName);
                    shape.ImageData.Save(bmpPath); // Save directly as BMP
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{docPath}'.");
        }

        // Validate that at least one BMP file was created
        int bmpCount = Directory.GetFiles(outputDir, "*.bmp").Length;
        if (bmpCount == 0)
            throw new InvalidOperationException("No BMP images were saved during batch processing.");

        Console.WriteLine($"Batch processing completed. Extracted {bmpCount} BMP image(s) to '{outputDir}'.");
    }

    private static void CreateSampleImage(string filePath)
    {
        const int width = 100;
        const int height = 100;

        // Create bitmap using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Black))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath);
        }
    }

    private static void CreateSampleDocuments(string folderPath, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample Document {i}");
            builder.InsertImage(imagePath);
            string docFileName = $"Doc{i}.docx";
            string docFullPath = Path.Combine(folderPath, docFileName);
            doc.Save(docFullPath);
        }
    }
}
