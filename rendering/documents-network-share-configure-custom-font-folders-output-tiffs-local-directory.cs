using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // UNC path to the folder that contains the source Word documents.
        string networkSharePath = @"\\Server\Share\Documents";

        // Local folder where the resulting TIFF files will be written.
        string tiffOutputPath = @"C:\Output\Tiffs";
        Directory.CreateDirectory(tiffOutputPath);

        // Folder(s) that contain custom TrueType fonts required for rendering.
        string[] customFontFolders = new[] { @"C:\CustomFonts" };

        // Configure Aspose.Words to search the custom font folders.
        // The second argument (true) enables recursive scanning of subfolders.
        FontSettings.DefaultInstance.SetFontsFolders(customFontFolders, true);

        // Process each .docx file found on the network share.
        foreach (string docFile in Directory.GetFiles(networkSharePath, "*.docx"))
        {
            // Load the document from the network location.
            Document doc = new Document(docFile);

            // Prepare TIFF save options. The constructor takes the desired SaveFormat.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

            // Example: render at 300 DPI for higher quality output.
            tiffOptions.Resolution = 300;

            // Build the full path for the output TIFF file.
            string outputFileName = Path.GetFileNameWithoutExtension(docFile) + ".tiff";
            string outputFilePath = Path.Combine(tiffOutputPath, outputFileName);

            // Save the entire document as a (multi‑page) TIFF image.
            doc.Save(outputFilePath, tiffOptions);
        }
    }
}
