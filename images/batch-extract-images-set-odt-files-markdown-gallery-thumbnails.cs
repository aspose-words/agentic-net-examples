using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OdtImageGallery
{
    // Custom callback to control image file names when saving to Markdown.
    class ImageRenamer : IImageSavingCallback
    {
        private readonly string _docBaseName;
        private int _imageIndex = 0;

        public ImageRenamer(string docBaseName) => _docBaseName = docBaseName;

        void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
        {
            // Preserve original extension, generate a deterministic name.
            string extension = Path.GetExtension(args.ImageFileName);
            string newFileName = $"{_docBaseName}_img{_imageIndex}{extension}";
            args.ImageFileName = newFileName;
            _imageIndex++;
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder containing the source ODT files (created if missing).
            string sourceFolder = Path.Combine(AppContext.BaseDirectory, "SourceOdtFiles");
            Directory.CreateDirectory(sourceFolder);

            // Folder where the Markdown files and extracted images will be placed.
            string outputFolder = Path.Combine(AppContext.BaseDirectory, "MarkdownGallery");
            Directory.CreateDirectory(outputFolder);

            // Process each ODT file in the source folder.
            foreach (string odtPath in Directory.GetFiles(sourceFolder, "*.odt"))
            {
                // Load the ODT document.
                Document doc = new Document(odtPath); // lifecycle: load

                // Base name without extension (used for naming output files).
                string baseName = Path.GetFileNameWithoutExtension(odtPath);

                // Create a sub‑folder for the images of this document.
                string imagesFolder = Path.Combine(outputFolder, $"{baseName}_images");
                Directory.CreateDirectory(imagesFolder);

                // Configure Markdown save options.
                var mdOptions = new MarkdownSaveOptions
                {
                    // Folder where Aspose.Words will write the extracted images.
                    ImagesFolder = imagesFolder,

                    // Optional: customize image file names.
                    ImageSavingCallback = new ImageRenamer(baseName)
                };

                // Path of the generated Markdown file.
                string markdownPath = Path.Combine(outputFolder, $"{baseName}.md");

                // Save the document as Markdown; this extracts images and creates the gallery.
                doc.Save(markdownPath, mdOptions); // lifecycle: save
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
