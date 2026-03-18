using System;
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
            // Load the existing DOCX document.
            Document doc = new Document(inputDocxPath);

            // Configure save options for MHTML.
            // - SaveFormat.Mhtml tells Aspose.Words to produce an MHTML (web archive) file.
            // - ExportFontResources = true ensures that font files are embedded.
            // - ExportImagesAsBase64 = true embeds images directly as Base64 data URIs.
            // - ExportCidUrlsForMhtmlResources = true uses CID URLs, which is the standard way
            //   to embed resources in MHTML files.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportFontResources = true,
                ExportImagesAsBase64 = true,
                ExportCidUrlsForMhtmlResources = true,
                // Optional: make the output more readable.
                PrettyFormat = true
            };

            // Save the document using the configured options.
            doc.Save(outputMhtmlPath, saveOptions);
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
            // Example usage – replace with real paths or pass via command‑line arguments.
            string inputPath = "input.docx";   // Path to the source DOCX file.
            string outputPath = "output.mhtml"; // Desired MHTML output path.

            // Call the conversion method.
            DocxToMhtmlConverter.Convert(inputPath, outputPath);

            Console.WriteLine($"DOCX converted to MHTML successfully: {outputPath}");
        }
    }
}
