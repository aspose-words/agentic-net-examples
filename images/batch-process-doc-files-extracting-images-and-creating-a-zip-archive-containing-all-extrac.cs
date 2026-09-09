using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, SolidBrush

public class Program
{
    public static void Main()
    {
        // Define working folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");

        // Prepare folders.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be inserted into documents.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to make the image recognizable.
                g.FillRectangle(new SolidBrush(Color.Blue), 10, 10, imgWidth - 20, imgHeight - 20);
            }
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Generate a few sample DOCX files that contain the image.
        // -----------------------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} – contains an image.");
            // Insert the previously created image.
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputDir, $"SampleDocument{i}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Batch process all DOC/DOCX files, extract images, and store them.
        // -----------------------------------------------------------------
        int extractedCount = 0;
        foreach (string docFile in Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Consider only Word document extensions.
            string ext = Path.GetExtension(docFile).ToLowerInvariant();
            if (ext != ".doc" && ext != ".docx") continue;

            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Build a deterministic file name: <DocName>_Img<index>.<extension>
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Img{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image data.
                shape.ImageData.Save(imagePath);
                extractedCount++;
                imageIndex++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -----------------------------------------------------------------
        // 4. Create a ZIP archive containing all extracted images.
        // -----------------------------------------------------------------
        if (File.Exists(zipPath)) File.Delete(zipPath);
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (string imageFile in Directory.GetFiles(outputDir))
            {
                string entryName = Path.GetFileName(imageFile);
                archive.CreateEntryFromFile(imageFile, entryName);
            }
        }

        // Validate that the ZIP file was created and contains entries.
        if (!File.Exists(zipPath))
            throw new FileNotFoundException("Failed to create the ZIP archive.");

        using (FileStream zipToRead = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive archive = new ZipArchive(zipToRead, ZipArchiveMode.Read))
        {
            if (archive.Entries.Count == 0)
                throw new InvalidOperationException("The ZIP archive is empty.");
        }

        Console.WriteLine($"Processed {docCount} document(s).");
        Console.WriteLine($"Extracted {extractedCount} image(s) to \"{outputDir}\".");
        Console.WriteLine($"Created ZIP archive: \"{zipPath}\".");
    }
}
