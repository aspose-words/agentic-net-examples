using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    /// <summary>
    /// Provides a utility to convert any supported document format to the legacy DOC format.
    /// </summary>
    public static class DocumentConverter
    {
        /// <summary>
        /// Converts the document read from <paramref name="inputStream"/> and writes the result in DOC format to <paramref name="outputStream"/>.
        /// </summary>
        /// <param name="inputStream">Stream containing the source document.</param>
        /// <param name="outputStream">Stream that will receive the converted DOC document.</param>
        public static void ConvertToDoc(Stream inputStream, Stream outputStream)
        {
            // Detect the format of the input document.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputStream);

            // Reset the stream position after detection.
            if (inputStream.CanSeek)
                inputStream.Position = 0;

            // Prepare load options based on the detected format.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = formatInfo.LoadFormat
            };

            // Load the document.
            Document doc = new Document(inputStream, loadOptions);

            // Configure save options for the legacy DOC format.
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Save the document to the output stream using the DOC save options.
            doc.Save(outputStream, saveOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: convert a file supplied via command‑line arguments.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: DocumentConverter <inputFile> <outputFile>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            using (FileStream inputStream = File.OpenRead(inputPath))
            using (FileStream outputStream = File.Create(outputPath))
            {
                DocumentConverter.ConvertToDoc(inputStream, outputStream);
            }

            Console.WriteLine($"Converted '{inputPath}' to DOC format as '{outputPath}'.");
        }
    }
}
