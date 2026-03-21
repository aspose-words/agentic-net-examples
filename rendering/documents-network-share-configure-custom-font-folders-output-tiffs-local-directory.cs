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

        // If the network share is unavailable, fall back to a local temporary folder.
        if (!Directory.Exists(networkSharePath))
        {
            Console.WriteLine("Network share not found. Falling back to a local folder.");
            networkSharePath = Path.Combine(Path.GetTempPath(), "Docs");
            Directory.CreateDirectory(networkSharePath);
        }

        // Local folder where the resulting TIFF files will be written.
        string tiffOutputPath = @"C:\Output\Tiffs";
        Directory.CreateDirectory(tiffOutputPath);

        // Folder(s) that contain custom TrueType fonts required for rendering.
        string[] customFontFolders = new[] { @"C:\CustomFonts" };

        // Configure Aspose.Words to search the custom font folders.
        // The second argument (true) enables recursive scanning of subfolders.
        FontSettings.DefaultInstance.SetFontsFolders(customFontFolders, true);

        // Process each .docx file found on the (network or fallback) share.
        string[] docFiles = Directory.GetFiles(networkSharePath, "*.docx", SearchOption.TopDirectoryOnly);
        if (docFiles.Length == 0)
        {
            Console.WriteLine($"No .docx files found in '{networkSharePath}'.");
            return;
        }

        foreach (string docFile in docFiles)
        {
            try
            {
                // Load the document from the location.
                Document doc = new Document(docFile);

                // Prepare TIFF save options.
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    Resolution = 300 // Render at 300 DPI for higher quality output.
                };

                // Build the full path for the output TIFF file.
                string outputFileName = Path.GetFileNameWithoutExtension(docFile) + ".tiff";
                string outputFilePath = Path.Combine(tiffOutputPath, outputFileName);

                // Save the entire document as a (multi‑page) TIFF image.
                doc.Save(outputFilePath, tiffOptions);
                Console.WriteLine($"Saved TIFF: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to process '{docFile}': {ex.Message}");
            }
        }
    }
}
