using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtraction");
        string docsDir = Path.Combine(baseDir, "Docs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean folders.
        foreach (string dir in new[] { docsDir, imagesDir, outputDir })
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, true);
            Directory.CreateDirectory(dir);
        }

        // -------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Simple deterministic content: a filled rectangle.
            graphics.FillRectangle(new SolidBrush(Aspose.Drawing.Color.LightBlue), 20, 20, 160, 160);
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a few sample DOCX files that contain the image.
        // -------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            // Insert the same image twice to have multiple shapes.
            builder.InsertImage(sampleImagePath);
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);

            string docPath = Path.Combine(docsDir, $"SampleDoc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images and build index data.
        // -------------------------------------------------
        var indexLines = new System.Collections.Generic.List<string>();
        // Header for CSV (Excel can open CSV directly).
        indexLines.Add("DocumentPath,ImageFileName,ImagePath");

        var docFiles = Directory.GetFiles(docsDir, "*.docx");
        int totalExtracted = 0;

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Determine proper file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);

                // Record index entry.
                string csvLine = $"\"{docFile}\",\"{imageFileName}\",\"{imageFullPath}\"";
                indexLines.Add(csvLine);

                imageIndex++;
                totalExtracted++;
            }
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -------------------------------------------------
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 5. Write the CSV index file (Excel-friendly).
        // -------------------------------------------------
        string indexCsvPath = Path.Combine(outputDir, "ImageIndex.csv");
        File.WriteAllLines(indexCsvPath, indexLines);

        // -------------------------------------------------
        // 6. Simple verification (optional).
        // -------------------------------------------------
        Console.WriteLine($"Processed {docFiles.Length} documents.");
        Console.WriteLine($"Extracted {totalExtracted} images to \"{imagesDir}\".");
        Console.WriteLine($"Index file created at \"{indexCsvPath}\".");
    }
}
