using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToMarkdownConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path to the resulting Markdown file.
        string markdownPath = @"C:\Output\sample.md";

        // Determine the folder where images will be stored (a subfolder named "assets").
        string imagesFolder = Path.Combine(Path.GetDirectoryName(markdownPath) ?? string.Empty, "assets");

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Load the PDF document.
        Document document = new Document(pdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Save images to the specified folder.
            ImagesFolder = imagesFolder,
            // Use a relative URI ("assets") in the Markdown file for image links.
            ImagesFolderAlias = "assets",
            // Explicitly set the format to Markdown (optional, but clarifies intent).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown, extracting images to the "assets" subfolder.
        document.Save(markdownPath, saveOptions);
    }
}
