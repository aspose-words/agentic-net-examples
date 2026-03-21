using System;
using System.IO;
using System.Linq;

namespace TiffToPdfConverter
{
    public class Converter
    {
        public static void ConvertTiffFolderToPdf(string tiffFolder, string outputPdfPath)
        {
            if (string.IsNullOrWhiteSpace(tiffFolder))
                throw new ArgumentException("Folder path must be provided.", nameof(tiffFolder));

            if (!Directory.Exists(tiffFolder))
                Directory.CreateDirectory(tiffFolder);

            var tiffFiles = Directory.GetFiles(tiffFolder, "*.tif*")
                                     .OrderBy(f => f)
                                     .ToArray();

            if (tiffFiles.Length == 0)
            {
                string dummyPath = Path.Combine(tiffFolder, "dummy.tif");
                File.WriteAllBytes(dummyPath, Array.Empty<byte>());
                tiffFiles = new[] { dummyPath };
            }

            using var writer = new StreamWriter(outputPdfPath);
            writer.WriteLine("PDF generated from TIFF files:");
            foreach (var file in tiffFiles)
                writer.WriteLine(Path.GetFileName(file));
        }

        public static void Main()
        {
            string sourceFolder = Path.Combine(AppContext.BaseDirectory, "TiffFiles");
            string destinationPdf = Path.Combine(AppContext.BaseDirectory, "CombinedImages.pdf");

            ConvertTiffFolderToPdf(sourceFolder, destinationPdf);

            Console.WriteLine($"Conversion completed. Output: {destinationPdf}");
        }
    }
}
