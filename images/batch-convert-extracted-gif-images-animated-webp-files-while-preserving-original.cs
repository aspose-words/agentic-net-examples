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
            // Use folders relative to the executable so they always exist.
            string baseDir = AppContext.BaseDirectory;
            string inputFolder = Path.Combine(baseDir, "Gif");
            string outputFolder = Path.Combine(baseDir, "WebP");

            // Ensure both folders exist.
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Get all GIF files in the input folder.
            string[] gifFiles = Directory.GetFiles(inputFolder, "*.gif", SearchOption.TopDirectoryOnly);

            if (gifFiles.Length == 0)
            {
                Console.WriteLine($"No GIF files found in \"{inputFolder}\". Place GIFs there and rerun.");
                return;
            }

            foreach (string gifPath in gifFiles)
            {
                // Load the GIF file as a document. Aspose.Words treats each frame of an animated GIF
                // as a separate page in the document, preserving the original frame delays.
                Document doc = new Document(gifPath);

                // Configure image save options for WebP format.
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.WebP);

                // Build the output file name with the same base name but .webp extension.
                string outputFileName = Path.GetFileNameWithoutExtension(gifPath) + ".webp";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the document as an animated WebP file.
                doc.Save(outputPath, saveOptions);
                Console.WriteLine($"Converted \"{Path.GetFileName(gifPath)}\" to \"{outputFileName}\".");
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
