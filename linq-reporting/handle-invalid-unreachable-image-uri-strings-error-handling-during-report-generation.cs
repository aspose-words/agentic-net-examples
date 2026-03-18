using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input and output paths.
        string inputPath = "Input.docx";
        string outputPath = "Report.html";

        // Folder where images will be saved and a local placeholder image.
        string imagesFolder = Path.Combine(Environment.CurrentDirectory, "Images");
        string placeholderPath = Path.Combine(Environment.CurrentDirectory, "placeholder.png");

        // Ensure the images folder exists.
        Directory.CreateDirectory(imagesFolder);

        // Load the document (create/load rule).
        LoadOptions loadOptions = new LoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Set callbacks to handle missing or unreachable images.
        doc.ResourceLoadingCallback = new ImageLoadingHandler(placeholderPath);
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ImagesFolder = imagesFolder,
            ImageSavingCallback = new ImageAvailabilityHandler(placeholderPath, imagesFolder)
        };

        // Build the report (example with an empty data source).
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages
        };
        engine.BuildReport(doc, new object(), string.Empty);

        // Save the document (save rule).
        doc.Save(outputPath, saveOptions);
    }
}

// Handles loading of external images. If the image cannot be downloaded,
// the placeholder image is supplied instead.
class ImageLoadingHandler : IResourceLoadingCallback
{
    private readonly string _placeholderPath;

    public ImageLoadingHandler(string placeholderPath)
    {
        _placeholderPath = placeholderPath;
    }

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        if (args.ResourceType != ResourceType.Image)
            return ResourceLoadingAction.Default;

        try
        {
            // Attempt to download the image from the original URI.
            using (HttpClient client = new HttpClient())
            {
                byte[] data = client.GetByteArrayAsync(args.OriginalUri).GetAwaiter().GetResult();
                args.SetData(data);
                return ResourceLoadingAction.UserProvided;
            }
        }
        catch
        {
            // On failure, use the local placeholder image if it exists.
            if (File.Exists(_placeholderPath))
            {
                args.SetData(File.ReadAllBytes(_placeholderPath));
                return ResourceLoadingAction.UserProvided;
            }

            // If no placeholder is available, skip the image.
            return ResourceLoadingAction.Skip;
        }
    }
}

// Handles image saving. If the original image is unavailable,
// the placeholder image data is written instead.
class ImageAvailabilityHandler : IImageSavingCallback
{
    private readonly string _placeholderPath;
    private readonly string _imagesFolder;

    public ImageAvailabilityHandler(string placeholderPath, string imagesFolder)
    {
        _placeholderPath = placeholderPath;
        _imagesFolder = imagesFolder;
    }

    void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
    {
        // Determine the target file path.
        string targetPath = Path.Combine(_imagesFolder, args.ImageFileName);

        if (!args.IsImageAvailable && File.Exists(_placeholderPath))
        {
            // Write placeholder image data to the target file.
            byte[] placeholderData = File.ReadAllBytes(_placeholderPath);
            args.ImageStream = new MemoryStream(placeholderData);
        }
        else
        {
            // Write the actual image data to the target file.
            args.ImageStream = new FileStream(targetPath, FileMode.Create);
        }

        // Ensure the stream is closed after Aspose.Words finishes writing.
        args.KeepImageStreamOpen = false;
    }
}
