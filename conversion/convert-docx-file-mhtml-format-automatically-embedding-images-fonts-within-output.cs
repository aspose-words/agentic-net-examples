using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    public class DocxToMhtmlConverter
    {
        /// <summary>
        /// Converts a DOCX file to MHTML, embedding all images and fonts into the resulting file.
        /// </summary>
        /// <param name="inputDocxPath">Full path to the source DOCX file.</param>
        /// <param name="outputMhtmlPath">Full path where the MHTML file will be saved.</param>
        public static void Convert(string inputDocxPath, string outputMhtmlPath)
        {
            if (!File.Exists(inputDocxPath))
            {
                Console.WriteLine($"Input file not found: {inputDocxPath}");
                return;
            }

            try
            {
                // Load the existing DOCX document.
                Document doc = new Document(inputDocxPath);

                // Configure save options for MHTML.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
                {
                    ExportFontResources = true,
                    ExportImagesAsBase64 = true,
                    ExportCidUrlsForMhtmlResources = true,
                    PrettyFormat = true
                };

                // Save the document using the configured options.
                doc.Save(outputMhtmlPath, saveOptions);
                Console.WriteLine($"DOCX converted to MHTML successfully: {outputMhtmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required by the C# compiler for a console application.
        /// Adjust the file paths as needed before running.
        /// </summary>
        static void Main(string[] args)
        {
            // Allow passing paths via command‑line arguments.
            string inputPath;
            string outputPath;

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }
            else
            {
                // Example usage – replace with real paths or pass via command‑line arguments.
                inputPath = "input.docx";   // Path to the source DOCX file.
                outputPath = "output.mhtml"; // Desired MHTML output path.
            }

            DocxToMhtmlConverter.Convert(inputPath, outputPath);
        }
    }
}
