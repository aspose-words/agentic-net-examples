using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path where the resulting Markdown file will be saved.
        string markdownPath = @"C:\Output\sample.md";

        // Create a temporary folder for extracted images.
        string imagesTempFolder = Path.Combine(Path.GetTempPath(),
            "AsposeImages_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(imagesTempFolder);

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Save images to the temporary folder.
            ImagesFolder = imagesTempFolder,
            // Use a dot as the folder alias so image URIs are relative (no path prefix).
            ImagesFolderAlias = ".",
            // Ensure the format is set to Markdown (optional, as the options type implies it).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown. Images will be written to the temporary folder,
        // and the Markdown file will reference them with relative paths.
        doc.Save(markdownPath, saveOptions);
    }
}
