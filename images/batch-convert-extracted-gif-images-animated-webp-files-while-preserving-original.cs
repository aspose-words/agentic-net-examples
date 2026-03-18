using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace GifToAnimatedWebP
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the extracted GIF files.
            string inputFolder = @"C:\Images\Gif";
            // Folder where the resulting animated WebP files will be saved.
            string outputFolder = @"C:\Images\WebP";

            Directory.CreateDirectory(outputFolder);

            // Process each GIF file in the input folder.
            foreach (string gifPath in Directory.GetFiles(inputFolder, "*.gif"))
            {
                // Load the GIF file as a document. Aspose.Words treats each frame of an animated GIF
                // as a separate page in the document, preserving the original frame delays.
                Document doc = new Document(gifPath);

                // Configure image save options for WebP format.
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.WebP)
                {
                    // For multi‑frame formats (GIF, WebP) the default layout renders each page as a separate frame.
                    // No additional layout configuration is required; the frame delays are kept automatically.
                    // If you need to customize the layout (e.g., grid), you could set:
                    // PageLayout = MultiPageLayout.Grid(1, 0, 0);
                };

                // Build the output file name with the same base name but .webp extension.
                string outputFileName = Path.GetFileNameWithoutExtension(gifPath) + ".webp";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the document as an animated WebP file.
                doc.Save(outputPath, saveOptions);
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
