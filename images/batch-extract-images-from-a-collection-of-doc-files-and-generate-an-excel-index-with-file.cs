using System;
using System.IO;
using System.Collections.Generic;
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
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string docsDir = Path.Combine(baseDir, "Docs");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string indexPath = Path.Combine(baseDir, "ImageIndex.xlsx"); // CSV with .xlsx extension.

        // Ensure clean folders.
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(extractedDir);
        if (File.Exists(indexPath)) File.Delete(indexPath);

        // -------------------------------------------------
        // 1. Create deterministic sample images.
        // -------------------------------------------------
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 150, 200, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -------------------------------------------------
        CreateDocumentWithImages(Path.Combine(docsDir, "Doc1.docx"), new[] { sampleImage1, sampleImage2 });
        CreateDocumentWithImages(Path.Combine(docsDir, "Doc2.docx"), new[] { sampleImage2 });

        // -------------------------------------------------
        // 3. Batch process all DOCX files: extract images and build index.
        // -------------------------------------------------
        var indexRows = new List<(string DocPath, string ImagePath)>();

        foreach (string docPath in Directory.GetFiles(docsDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Collect all shapes that actually contain images.
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in document '{docPath}'.");

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine proper file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(extractedDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);

                indexRows.Add((docPath, imageFullPath));
                imageIndex++;
            }
        }

        if (indexRows.Count == 0)
            throw new InvalidOperationException("No images were extracted from any document.");

        // -------------------------------------------------
        // 4. Write the index to a CSV file (named .xlsx for Excel compatibility).
        // -------------------------------------------------
        using (var writer = new StreamWriter(indexPath, false))
        {
            writer.WriteLine("DocumentPath,ImagePath");
            foreach (var row in indexRows)
            {
                // Escape commas if they ever appear in paths.
                string docEscaped = EscapeCsv(row.DocPath);
                string imgEscaped = EscapeCsv(row.ImagePath);
                writer.WriteLine($"{docEscaped},{imgEscaped}");
            }
        }

        // Validation: ensure the index file exists and contains at least one line besides header.
        if (!File.Exists(indexPath) || new FileInfo(indexPath).Length == 0)
            throw new InvalidOperationException("Failed to create the Excel index file.");

        // Example completed without interactive prompts.
    }

    // Creates a simple bitmap image with a solid background color.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (var bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file and inserts the provided image files.
    private static void CreateDocumentWithImages(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imagePaths)
        {
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");

            // Insert the image inline.
            builder.InsertImage(imgPath);
            builder.Writeln(); // Add a line break between images.
        }

        doc.Save(docPath);
    }

    // Simple CSV escaping for fields that may contain commas or quotes.
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
