using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToTiffBatch
{
    class Program
    {
        /// <summary>
        /// Converts all DOCX files in the specified folder to multi‑page TIFF files.
        /// The TIFF compression scheme is set via the ImageSaveOptions.TiffCompression property.
        /// </summary>
        /// <param name="args">
        /// args[0] (optional) – Path to the folder containing DOCX files.
        /// If omitted, the current working directory is used.
        /// </param>
        static void Main(string[] args)
        {
            // Determine the source folder.
            string sourceFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Create an output sub‑folder for the generated TIFF files.
            string outputFolder = Path.Combine(sourceFolder, "TiffOutput");
            Directory.CreateDirectory(outputFolder);

            // Enumerate all DOCX files in the source folder.
            foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
            {
                // Load the DOCX document.
                Document doc = new Document(docxPath);

                // Configure image save options for TIFF output.
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    // Use LZW compression (change to desired enum value if needed).
                    TiffCompression = TiffCompression.Lzw,

                    // Render all pages into a single multi‑frame TIFF.
                    PageLayout = MultiPageLayout.TiffFrames()
                };

                // Build the output file name (same base name, .tiff extension).
                string tiffPath = Path.Combine(outputFolder,
                    Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

                // Save the document as TIFF using the configured options.
                doc.Save(tiffPath, tiffOptions);
            }

            Console.WriteLine("Conversion completed. TIFF files are located in:");
            Console.WriteLine(outputFolder);
        }
    }
}
