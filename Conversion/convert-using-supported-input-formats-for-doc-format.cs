using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace ConvertToDocExample
{
    public class DocConverter
    {
        /// <summary>
        /// Converts a document from any supported input format to the legacy DOC format.
        /// </summary>
        /// <param name="inputPath">Full path to the source document.</param>
        /// <param name="outputPath">Full path where the DOC file will be saved.</param>
        public void ConvertToDoc(string inputPath, string outputPath)
        {
            // Detect the format of the input file.
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);

            // If the format cannot be detected or is unknown, abort the conversion.
            if (formatInfo.LoadFormat == LoadFormat.Unknown)
                throw new InvalidOperationException("The input file format is not supported.");

            // Create LoadOptions with the detected format to ensure proper loading.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = formatInfo.LoadFormat
            };

            // Load the document using the determined options.
            Document doc = new Document(inputPath, loadOptions);

            // Prepare save options for the DOC format.
            DocSaveOptions saveOptions = new DocSaveOptions
            {
                SaveFormat = SaveFormat.Doc // Explicitly set the target format.
            };

            // Save the document as DOC.
            doc.Save(outputPath, saveOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Simple argument handling: first argument = input file, second argument = output file.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ConvertToDocExample <inputPath> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                var converter = new DocConverter();
                converter.ConvertToDoc(inputPath, outputPath);
                Console.WriteLine($"Successfully converted '{inputPath}' to DOC at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
