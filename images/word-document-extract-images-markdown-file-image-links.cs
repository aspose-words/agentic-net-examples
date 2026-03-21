using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractImagesToMarkdown
{
    static void Main()
    {
        // Create a temporary working folder.
        string outputRoot = Path.Combine(Path.GetTempPath(), "AsposeExtractDemo");
        Directory.CreateDirectory(outputRoot);

        // Path for the sample Word document.
        string sourceDocPath = Path.Combine(outputRoot, "Sample.docx");

        // If the sample document does not exist, create it with a simple image.
        if (!File.Exists(sourceDocPath))
        {
            // Create a tiny PNG image (1x1 pixel, red) from a Base64 string.
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            string tempImagePath = Path.Combine(outputRoot, "red.png");
            File.WriteAllBytes(tempImagePath, pngBytes);

            // Build the Word document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document with an image:");
            builder.InsertImage(tempImagePath);
            doc.Save(sourceDocPath);
        }

        // Folder that will hold the extracted image files.
        string imagesFolder = Path.Combine(outputRoot, "images");
        Directory.CreateDirectory(imagesFolder);

        // Full path for the generated markdown file.
        string markdownPath = Path.Combine(outputRoot, "Sample.md");

        // Load the Word document.
        Document sourceDoc = new Document(sourceDocPath);

        // Configure markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Physical location for extracted images.
            ImagesFolderAlias = "images",         // Relative URI used in markdown.
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as markdown; images are automatically extracted.
        sourceDoc.Save(markdownPath, saveOptions);

        Console.WriteLine($"Markdown file created at: {markdownPath}");
        Console.WriteLine($"Extracted images folder: {imagesFolder}");
    }
}
