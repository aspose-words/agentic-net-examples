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
            // Path to the source DOCX file.
            string inputDocxPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting Markdown file will be saved.
            string outputMarkdownPath = @"C:\Docs\ResultDocument.md";

            // Folder where extracted images will be stored.
            string imagesFolderPath = @"C:\Docs\ExtractedImages";

            // Ensure the images folder exists.
            Directory.CreateDirectory(imagesFolderPath);

            // Load the DOCX document from the file system.
            Document doc = new Document(inputDocxPath);

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Explicitly set the format to Markdown (optional, but clarifies intent).
                SaveFormat = SaveFormat.Markdown,

                // Specify the folder where images will be saved.
                ImagesFolder = imagesFolderPath,

                // Optionally, set an alias for image URIs in the Markdown file.
                // ImagesFolderAlias = "images"
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputMarkdownPath, saveOptions);
        }
    }
}
