using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading; // Added for LoadOptions
using Aspose.Words.Drawing;

class ExtractVectorImagesToEmf
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\source.doc";

        // Folder where the extracted EMF files will be saved.
        string outputFolder = @"C:\Docs\ExtractedEmf";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document without converting metafiles to raster images.
        LoadOptions loadOptions = new LoadOptions();
        // Explicitly keep vector metafiles (WMF/EMF) as they are.
        loadOptions.ConvertMetafilesToPng = false;

        Document doc = new Document(inputPath, loadOptions);

        int imageCounter = 0;

        // Iterate over all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Skip shapes that do not contain an image.
            if (!shape.HasImage) continue;

            // Identify vector metafile images (WMF or EMF).
            ImageType imgType = shape.ImageData.ImageType;
            if (imgType == ImageType.Wmf || imgType == ImageType.Emf)
            {
                // Build a unique file name for each extracted image.
                string outFile = Path.Combine(outputFolder, $"vector_{imageCounter}.emf");

                // Save the image. The file extension determines the output format.
                // If the source is WMF, Aspose.Words will convert it to EMF automatically.
                shape.ImageData.Save(outFile);

                imageCounter++;
            }
        }

        Console.WriteLine($"Extracted {imageCounter} vector image(s) to EMF format.");
    }
}
