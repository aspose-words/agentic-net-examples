using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageBatchProcessor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(baseDir);

        // Directories for documents and extracted images.
        string docsDir = Path.Combine(baseDir, "Docs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);

        // Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Create a few sample DOCX files each containing the sample image.
        CreateSampleDocuments(docsDir, sampleImagePath, 3);

        // List to hold CSV rows.
        List<string[]> csvRows = new List<string[]>();
        // Header row.
        csvRows.Add(new[] { "DocumentName", "ImageFileName", "ImageType", "WidthPoints", "HeightPoints", "ImageSizeBytes" });

        // Process each DOCX file.
        foreach (string docPath in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save image: {imageFullPath}");

                // Gather metadata.
                long sizeBytes = shape.ImageData.ImageBytes?.Length ?? shape.ImageData.ToByteArray().Length;
                string[] row = new[]
                {
                    Path.GetFileName(docPath),
                    imageFileName,
                    shape.ImageData.ImageType.ToString(),
                    shape.Width.ToString(),
                    shape.Height.ToString(),
                    sizeBytes.ToString()
                };
                csvRows.Add(row);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (csvRows.Count <= 1)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Write CSV report.
        string csvPath = Path.Combine(baseDir, "ImageReport.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            foreach (var row in csvRows)
            {
                writer.WriteLine(string.Join(",", row.Select(EscapeCsv)));
            }
        }

        // Completion messages.
        Console.WriteLine($"Processing complete. Extracted images are in: {imagesDir}");
        Console.WriteLine($"CSV report generated at: {csvPath}");
    }

    // Creates a simple white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Generates a number of DOCX files each containing the sample image.
    private static void CreateSampleDocuments(string docsDir, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the sample image twice per document.
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
            builder.InsertImage(imagePath);

            string docFileName = Path.Combine(docsDir, $"SampleDoc{i}.docx");
            doc.Save(docFileName);
        }
    }

    // Escapes CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
