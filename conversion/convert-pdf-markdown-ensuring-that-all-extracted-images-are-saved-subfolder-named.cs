using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToMarkdownConverter
{
    static void Main()
    {
        // Use the system temporary directory for all files.
        string tempDir = Path.GetTempPath();

        // Paths for the sample PDF and the resulting Markdown file.
        string pdfPath = Path.Combine(tempDir, "sample.pdf");
        string markdownPath = Path.Combine(tempDir, "sample.md");

        // Ensure a sample PDF exists. If not, create a simple one.
        if (!File.Exists(pdfPath))
        {
            // Create a simple Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, world!");
            builder.InsertImage(ImageFromBase64()); // Insert a sample image.
            // Save it as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Determine the folder where images will be stored (a subfolder named "assets").
        string imagesFolder = Path.Combine(Path.GetDirectoryName(markdownPath) ?? string.Empty, "assets");

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Load the PDF document.
        Document document = new Document(pdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesFolder,          // Save images to the specified folder.
            ImagesFolderAlias = "assets",         // Use a relative URI ("assets") in the Markdown file for image links.
            SaveFormat = SaveFormat.Markdown     // Explicitly set the format to Markdown.
        };

        // Save the document as Markdown, extracting images to the "assets" subfolder.
        document.Save(markdownPath, saveOptions);

        Console.WriteLine($"Markdown saved to: {markdownPath}");
        Console.WriteLine($"Images saved to: {imagesFolder}");
    }

    // Helper method to provide a small PNG image from a Base64 string.
    private static Stream ImageFromBase64()
    {
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4" +
            "//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==";
        byte[] bytes = Convert.FromBase64String(base64Png);
        return new MemoryStream(bytes);
    }
}
