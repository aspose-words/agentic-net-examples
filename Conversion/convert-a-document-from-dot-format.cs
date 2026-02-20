using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace ConvertDotExample
{
    public class DotConverter
    {
        /// <summary>
        /// Converts a Microsoft Word DOT template to DOCX format.
        /// </summary>
        /// <param name="inputStream">Stream containing the source DOT file.</param>
        /// <param name="outputStream">Stream where the converted DOCX will be written.</param>
        public void ConvertDotToDocx(Stream inputStream, Stream outputStream)
        {
            // Specify that the input format is a DOT template.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Dot
            };

            // Load the DOT document from the input stream.
            Document document = new Document(inputStream, loadOptions);

            // Prepare save options for DOCX output.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

            // Save the document to the output stream in DOCX format.
            document.Save(outputStream, saveOptions);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Simple demo: convert "template.dot" to "output.docx" located in the same folder as the executable.
            string inputPath = Path.Combine(AppContext.BaseDirectory, "template.dot");
            string outputPath = Path.Combine(AppContext.BaseDirectory, "output.docx");

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            using (FileStream inputStream = File.OpenRead(inputPath))
            using (FileStream outputStream = File.Create(outputPath))
            {
                DotConverter converter = new DotConverter();
                converter.ConvertDotToDocx(inputStream, outputStream);
            }

            Console.WriteLine($"Conversion completed. DOCX saved to: {outputPath}");
        }
    }
}
