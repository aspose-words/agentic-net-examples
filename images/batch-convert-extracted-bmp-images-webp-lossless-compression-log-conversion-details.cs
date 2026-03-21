#nullable disable
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
            // Ensure the log directory exists before any file operations.
            string logDir = Path.GetDirectoryName(logFilePath);
            if (!string.IsNullOrEmpty(logDir))
                Directory.CreateDirectory(logDir);

            // Ensure input folder exists; if not, log and exit gracefully.
            if (!Directory.Exists(inputFolder))
            {
                Console.Error.WriteLine($"Input folder not found: {inputFolder}");
                using (var logWriter = new StreamWriter(logFilePath, false))
                {
                    logWriter.WriteLine($"Input folder not found: {inputFolder}");
                }
                return;
            }

            Directory.CreateDirectory(outputFolder);

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
                        if (originalSize > 0)
                        {
                            double reduction = (originalSize - newSize) * 100.0 / originalSize;
                            logWriter.WriteLine($"  Reduction     : {reduction:F2}%");
                        }
                        else
                        {
                            logWriter.WriteLine("  Reduction     : N/A (original size zero)");
                        }
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
        /// If omitted, example paths relative to the current directory are used.
        /// </summary>
        static void Main(string[] args)
        {
            string baseDir = AppContext.BaseDirectory;
            string defaultInput = Path.Combine(baseDir, "Images", "Bmp");
            string defaultOutput = Path.Combine(baseDir, "Images", "Webp");
            string defaultLog = Path.Combine(baseDir, "Images", "conversion_log.txt");

            string inputFolder = args.Length > 0 ? args[0] : defaultInput;
            string outputFolder = args.Length > 1 ? args[1] : defaultOutput;
            string logFilePath = args.Length > 2 ? args[2] : defaultLog;

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
