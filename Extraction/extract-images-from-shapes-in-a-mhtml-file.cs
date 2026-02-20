using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImageExtractor : IImageSavingCallback
{
    public void ImageSaving(ImageSavingArgs args)
    {
        // Create a folder for the extracted images if it does not exist.
        const string outputFolder = "ExtractedImages";
        Directory.CreateDirectory(outputFolder);

        // Build the full path for the image file.
        string imagePath = Path.Combine(outputFolder, args.ImageFileName);

        // Write the image stream to the file.
        using (FileStream fileStream = new FileStream(imagePath, FileMode.Create))
        {
            args.ImageStream.CopyTo(fileStream);
        }

        // Instruct Aspose.Words not to keep the stream open after the callback.
        args.KeepImageStreamOpen = false;
    }
}

class Program
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("input.mht");

        // Configure save options to trigger the ImageSaving callback.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = false,          // Save images as separate files.
            ImageSavingCallback = new ImageExtractor()
        };

        // Save to a temporary file; the callback extracts all images.
        doc.Save("temp.mht", saveOptions);
    }
}
