using System;
using System.IO;
using Aspose.Words;

namespace DocumentConversionExample
{
    public class Converter
    {
        /// <summary>
        /// Converts a document from one format to another.
        /// The input and output formats are inferred from the file extensions.
        /// </summary>
        /// <param name="inputPath">Full path to the source document.</param>
        /// <param name="outputPath">Full path where the converted document will be saved.</param>
        public void Convert(string inputPath, string outputPath)
        {
            // Load the source document. The constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Determine the target save format from the output file extension.
            SaveFormat targetFormat = GetSaveFormatFromExtension(Path.GetExtension(outputPath));

            // Save the document in the desired format.
            doc.Save(outputPath, targetFormat);
        }

        /// <summary>
        /// Maps a file extension to the corresponding Aspose.Words SaveFormat value.
        /// </summary>
        private SaveFormat GetSaveFormatFromExtension(string extension)
        {
            // Ensure the extension starts with a dot and is in lower case.
            string ext = extension.StartsWith(".") ? extension.ToLowerInvariant() : "." + extension.ToLowerInvariant();

            // Use Aspose.Words utility to convert the extension to a SaveFormat.
            // This method throws if the extension is not supported.
            return FileFormatUtil.ExtensionToSaveFormat(ext);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            var converter = new Converter();

            // Example: convert a DOCX file to PDF.
            string inputFile = @"C:\Docs\SampleDocument.docx";
            string outputFile = @"C:\Docs\SampleDocument.pdf";

            converter.Convert(inputFile, outputFile);

            Console.WriteLine("Conversion completed.");
        }
    }
}
