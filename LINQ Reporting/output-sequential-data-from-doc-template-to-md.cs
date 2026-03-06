using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsToMarkdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOC template.
            // Replace with the actual file location.
            string templatePath = @"C:\Docs\Template.docx";

            // Path where the generated Markdown file will be saved.
            // Replace with the desired output location.
            string markdownPath = @"C:\Docs\Result.md";

            // Load the DOC template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Configure Markdown save options if any custom behavior is required.
            // Here we use the default options; you can modify properties such as
            // ExportImagesAsBase64, ImagesFolder, etc., as needed.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                // Example: Save images as Base64 strings inside the Markdown file.
                // ExportImagesAsBase64 = true,

                // Example: Specify a folder for extracted images (if not using Base64).
                // ImagesFolder = @"C:\Docs\Images"
            };

            // Save the document as Markdown using the configured options.
            doc.Save(markdownPath, mdOptions);
        }
    }
}
