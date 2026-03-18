using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ImageConversionDemo
{
    /// <summary>
    /// Provides functionality to batch‑convert BMP images to lossless WebP using Aspose.Words.
    /// </summary>
    public static class BmpToWebpConverter
    {
        /// <summary>
        /// Converts all *.bmp files found in <paramref name="inputFolder"/> to WebP format,
        /// stores the results in <paramref name="outputFolder"/> and writes a conversion log
        /// to <paramref name="logFilePath"/>.
        /// </summary>
        public static void Convert(string inputFolder, string outputFolder, string logFilePath)
        {
            // Ensure folders exist.
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            Directory.CreateDirectory(outputFolder);

            // Guard against Path.GetDirectoryName returning null (e.g., when logFilePath is a root path).
            string? logDir = Path.GetDirectoryName(logFilePath);
            if (!string.IsNullOrEmpty(logDir))
                Directory.CreateDirectory(logDir);

            // Prepare a log writer.
            using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
            {
                // Write header.
                logWriter.WriteLine("BMP to WebP Conversion Log");
                logWriter.WriteLine($"Start Time: {DateTime.Now}");
                logWriter.WriteLine($"Input Folder : {inputFolder}");
                logWriter.WriteLine($"Output Folder: {outputFolder}");
                logWriter.WriteLine(new string('-', 80));

                // Process each BMP file.
                foreach (string bmpPath in Directory.GetFiles(inputFolder, "*.bmp", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        long originalSize = new FileInfo(bmpPath).Length;

                        // Create a new empty document and insert the BMP image.
                        Document doc = new Document();
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        builder.InsertImage(bmpPath);

                        // Save options for lossless WebP (default is lossless).
                        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.WebP);

                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpPath);
                        string webpPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.webp");

                        // Save the document (which contains only the image) as WebP.
                        doc.Save(webpPath, saveOptions);

                        long newSize = new FileInfo(webpPath).Length;

                        // Log conversion details.
                        logWriter.WriteLine($"Converted: {Path.GetFileName(bmpPath)}");
                        logWriter.WriteLine($"  Original Size : {originalSize:N0} bytes");
                        logWriter.WriteLine($"  WebP Size     : {newSize:N0} bytes");
                        logWriter.WriteLine($"  Reduction     : {((originalSize - newSize) * 100.0 / originalSize):F2}%");
                        logWriter.WriteLine();
                    }
                    catch (Exception ex)
                    {
                        logWriter.WriteLine($"Failed to convert: {Path.GetFileName(bmpPath)}");
                        logWriter.WriteLine($"  Error: {ex.Message}");
                        logWriter.WriteLine();
                    }
                }

                // Write footer.
                logWriter.WriteLine(new string('-', 80));
                logWriter.WriteLine($"End Time: {DateTime.Now}");
            }
        }
    }

    /// <summary>
    /// Entry point for the console application.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Expected arguments: <c>inputFolder outputFolder logFilePath</c>.
        /// If omitted, example paths are used.
        /// </summary>
        static void Main(string[] args)
        {
            string inputFolder = args.Length > 0 ? args[0] : @"C:\Images\Bmp";
            string outputFolder = args.Length > 1 ? args[1] : @"C:\Images\Webp";
            string logFilePath = args.Length > 2 ? args[2] : @"C:\Images\conversion_log.txt";

            try
            {
                BmpToWebpConverter.Convert(inputFolder, outputFolder, logFilePath);
                Console.WriteLine("Conversion completed. See log for details.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
