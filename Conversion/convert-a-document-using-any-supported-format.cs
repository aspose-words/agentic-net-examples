using System;
using Aspose.Words;

namespace DocumentConversionExample
{
    public class Converter
    {
        /// <summary>
        /// Converts a document from its current format to the specified target format.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source document.</param>
        /// <param name="outputFilePath">Full path where the converted document will be saved.</param>
        /// <param name="targetFormat">The Aspose.Words.SaveFormat value representing the desired output format.</param>
        public static void Convert(string inputFilePath, string outputFilePath, SaveFormat targetFormat)
        {
            // Load the source document. The constructor automatically detects the format.
            Document doc = new Document(inputFilePath);

            // Save the document in the requested format.
            // The overload (string, SaveFormat) is part of the Aspose.Words API.
            doc.Save(outputFilePath, targetFormat);
        }

        // Example usage
        public static void Main()
        {
            // Convert a DOCX file to PDF.
            string sourcePath = @"C:\Docs\SampleDocument.docx";
            string destinationPath = @"C:\Docs\SampleDocument.pdf";

            Convert(sourcePath, destinationPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
