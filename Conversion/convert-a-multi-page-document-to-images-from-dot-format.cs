using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Callback to save each page of the document as a separate image file.
    public class ImagePageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _imageExtension;

        public ImagePageSavingCallback(string outputFolder, string imageExtension)
        {
            _outputFolder = outputFolder;
            _imageExtension = imageExtension;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Build the file name for the current page.
            string fileName = Path.Combine(_outputFolder, $"Page_{args.PageIndex}{_imageExtension}");

            // Set the file name – Aspose.Words will create the file automatically.
            args.PageFileName = fileName;

            // Alternatively, you could provide a custom stream:
            // args.PageStream = new FileStream(fileName, FileMode.Create);
        }
    }

    public class ConvertDotToImages
    {
        public static void Main()
        {
            // Path to the source DOT (Word template) file.
            string dotFilePath = @"C:\Docs\Template.dot";

            // Folder where the resulting images will be saved.
            string outputFolder = @"C:\Docs\Images";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the DOT document.
            Document doc = new Document(dotFilePath);

            // Configure image save options.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render all pages.
                PageSet = PageSet.All,

                // Optional: set resolution (dpi) for higher quality.
                HorizontalResolution = 300,
                VerticalResolution = 300,

                // Use a callback to name each page image file.
                PageSavingCallback = new ImagePageSavingCallback(outputFolder, ".png")
            };

            // The file name passed to Save is not used because the callback provides per‑page names.
            // It can be any valid path; we use a dummy name.
            doc.Save(Path.Combine(outputFolder, "dummy.png"), saveOptions);
        }
    }
}
