using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class ExtractImagesToZip
{
    // Entry point of the console application.
    public static void Main()
    {
        // Base directories for sample documents, extracted images and the final ZIP archive.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExampleData");
        string docsDir = Path.Combine(baseDir, "Docs");
        string imagesExtractDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ExtractedImages.zip");

        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesExtractDir);

        // Create sample documents containing images.
        CreateSampleDocuments(docsDir);

        // Extract images from all documents in the collection.
        int totalExtracted = ExtractImagesFromDocuments(docsDir, imagesExtractDir);

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the document collection.");

        // Create a ZIP archive containing all extracted images.
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(imagesExtractDir, zipPath);

        // Validation: ensure the ZIP file was created and contains the expected number of entries.
        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
        {
            if (archive.Entries.Count != totalExtracted)
                throw new InvalidOperationException("The number of files in the ZIP archive does not match the extracted image count.");
        }

        // The example finishes without requiring any user interaction.
    }

    // Generates a few sample DOCX files each containing an embedded image.
    private static void CreateSampleDocuments(string outputFolder)
    {
        // A tiny 1x1 PNG (red pixel) encoded in base64.
        const string redPixelBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "b9VYAAAAASUVORK5CYII=";

        byte[] imageBytes = Convert.FromBase64String(redPixelBase64);

        // Create two documents; each will receive the same image.
        for (int docIndex = 0; docIndex < 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph of text.
            builder.Writeln($"Sample document {docIndex + 1}");

            // Insert the image from a memory stream.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                // InsertImage returns the created Shape containing the image.
                Shape imageShape = builder.InsertImage(ms);
                // Ensure the shape is treated as an image shape.
                imageShape.WrapType = WrapType.Inline;
            }

            // Save the document.
            string docPath = Path.Combine(outputFolder, $"SampleDocument{docIndex + 1}.docx");
            doc.Save(docPath);
        }
    }

    // Traverses all DOCX files in the specified folder, extracts images from shape nodes,
    // and saves them into the target folder. Returns the total number of extracted images.
    private static int ExtractImagesFromDocuments(string docsFolder, string outputFolder)
    {
        int extractedCount = 0;
        string[] docFiles = Directory.GetFiles(docsFolder, "*.docx", SearchOption.TopDirectoryOnly);

        for (int docIdx = 0; docIdx < docFiles.Length; docIdx++)
        {
            Document doc = new Document(docFiles[docIdx]);

            // Retrieve all Shape nodes (including those inside headers/footers).
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .Where(s => s.HasImage)
                            .ToArray();

            for (int imgIdx = 0; imgIdx < shapes.Length; imgIdx++)
            {
                Shape shape = shapes[imgIdx];
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Deterministic file name: DocumentIndex_ImageIndex.extension
                string fileName = $"Doc{docIdx + 1}_Img{imgIdx + 1}{extension}";
                string filePath = Path.Combine(outputFolder, fileName);
                shape.ImageData.Save(filePath);
                extractedCount++;
            }
        }

        return extractedCount;
    }
}
