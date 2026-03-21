using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExample
{
    class Program
    {
        static void Main()
        {
            // Create a temporary working directory.
            string workDir = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
            Directory.CreateDirectory(workDir);

            // Define paths for the source DOCX, the resulting Markdown, and the images folder.
            string inputDocxPath = Path.Combine(workDir, "SourceDocument.docx");
            string outputMarkdownPath = Path.Combine(workDir, "ResultDocument.md");
            string imagesFolderPath = Path.Combine(workDir, "ExtractedImages");

            // Ensure the images folder exists.
            Directory.CreateDirectory(imagesFolderPath);

            // -----------------------------------------------------------------
            // Create a sample DOCX document with some text and an embedded image.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, this is a sample document generated at runtime.");

            // Insert a simple 1x1 PNG image from a base64 string (no System.Drawing dependency).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imgStream);
            }

            // Save the DOCX to the temporary location.
            doc.Save(inputDocxPath);

            // --------------------------------------------------------------
            // Load the DOCX and convert it to Markdown, extracting images.
            // --------------------------------------------------------------
            Document loadedDoc = new Document(inputDocxPath);
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown,
                ImagesFolder = imagesFolderPath,
                // Optionally set a relative path alias for images in the markdown.
                ImagesFolderAlias = "ExtractedImages"
            };

            loadedDoc.Save(outputMarkdownPath, saveOptions);

            Console.WriteLine("Conversion completed.");
            Console.WriteLine($"Source DOCX: {inputDocxPath}");
            Console.WriteLine($"Markdown output: {outputMarkdownPath}");
            Console.WriteLine($"Extracted images folder: {imagesFolderPath}");
        }
    }
}
