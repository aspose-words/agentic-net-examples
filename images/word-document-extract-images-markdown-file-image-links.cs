using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractImagesToMarkdown
{
    static void Main()
    {
        // Path to the source Word document.
        string sourceDocPath = @"C:\Docs\SourceDocument.docx";

        // Folder where the markdown file and extracted images will be placed.
        string outputRoot = @"C:\Docs\Output";
        Directory.CreateDirectory(outputRoot);

        // Folder that will hold the extracted image files.
        string imagesFolder = Path.Combine(outputRoot, "images");
        Directory.CreateDirectory(imagesFolder);

        // Full path for the generated markdown file.
        string markdownPath = Path.Combine(outputRoot, "SourceDocument.md");

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Configure markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Physical location where Aspose.Words will write the image files.
            ImagesFolder = imagesFolder,

            // URI that will be written into the markdown file for each image.
            // Using a relative folder name makes the markdown portable.
            ImagesFolderAlias = "images",

            // Explicitly set the format to Markdown (optional, but clear).
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as markdown; images are automatically extracted to ImagesFolder
        // and markdown links are created using ImagesFolderAlias.
        doc.Save(markdownPath, saveOptions);
    }
}
