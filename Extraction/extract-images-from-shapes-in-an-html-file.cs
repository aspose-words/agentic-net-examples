using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImageExtractor : IImageSavingCallback
{
    private readonly string _outputFolder;
    private int _imageIndex = 0;

    public ImageExtractor(string outputFolder)
    {
        _outputFolder = outputFolder;
        Directory.CreateDirectory(_outputFolder);
    }

    void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
    {
        // Only process images that are actually available.
        if (!args.IsImageAvailable)
            return;

        // Determine a unique file name for each extracted image.
        string extension = Path.GetExtension(args.ImageFileName);
        if (string.IsNullOrEmpty(extension))
            extension = ".png"; // fallback

        string fileName = $"image_{++_imageIndex}{extension}";
        string fullPath = Path.Combine(_outputFolder, fileName);

        // Direct Aspose.Words to write the image to the specified file.
        args.ImageStream = new FileStream(fullPath, FileMode.Create);
        args.KeepImageStreamOpen = false; // let Aspose close the stream after saving.
    }
}

class Program
{
    static void Main()
    {
        // Path to the source HTML file.
        string htmlPath = @"C:\Input\sample.html";

        // Folder where extracted images will be saved.
        string imagesFolder = @"C:\Output\ExtractedImages";

        // Load the HTML document with default options.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Document doc = new Document(htmlPath, loadOptions);

        // Configure save options with the custom image‑saving callback.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ImageSavingCallback = new ImageExtractor(imagesFolder);

        // Save the document back to HTML (the content itself is irrelevant;
        // the callback will be invoked for each image and write them to disk).
        string dummyOutput = Path.Combine(imagesFolder, "dummy.html");
        doc.Save(dummyOutput, saveOptions);
    }
}
