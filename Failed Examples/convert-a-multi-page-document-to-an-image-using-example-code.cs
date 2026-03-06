// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocumentToImages
{
    static void Main()
    {
        // Path to the source multi‑page document.
        const string inputPath = @"C:\Docs\MultiPageDocument.docx";

        // Folder where the resulting images will be saved.
        const string outputFolder = @"C:\Docs\Images";
        Directory.CreateDirectory(outputFolder);

        // Open the source document as a read‑only stream.
        using (FileStream inputStream = File.OpenRead(inputPath))
        {
            // Load options – default settings are sufficient for most cases.
            LoadOptions loadOptions = new LoadOptions();

            // Save options – configure the image format and rendering settings.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render all pages of the document.
                PageSet = PageSet.All,

                // Example: increase resolution for higher‑quality output.
                Resolution = 300
            };

            // Create an instance of the converter plugin (implementation provided by Aspose.Words).
            IDocumentConverterPlugin converter = new DocumentConverterPlugin();

            // Convert each page of the document to an image stream.
            Stream[] imageStreams = converter.ConvertToImages(inputStream, loadOptions, saveOptions);

            // Save each image stream to a separate file.
            for (int i = 0; i < imageStreams.Length; i++)
            {
                // Ensure the stream position is at the beginning before reading.
                imageStreams[i].Position = 0;

                string outputPath = Path.Combine(outputFolder, $"Page_{i + 1}.jpg");
                using (FileStream outputFile = File.Create(outputPath))
                {
                    imageStreams[i].CopyTo(outputFile);
                }

                // Dispose the individual image stream after saving.
                imageStreams[i].Dispose();
            }
        }
    }
}
