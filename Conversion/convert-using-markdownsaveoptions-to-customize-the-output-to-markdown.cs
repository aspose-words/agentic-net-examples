using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    public class DocmToMarkdownConverter
    {
        /// <summary>
        /// Converts a DOCM file to Markdown format with customized save options.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOCM file.</param>
        /// <param name="outputPath">Full path where the Markdown file will be saved.</param>
        public static void Convert(string inputPath, string outputPath)
        {
            // Load the DOCM document from the specified file.
            Document doc = new Document(inputPath);

            // Create MarkdownSaveOptions to customize the conversion.
            var saveOptions = new MarkdownSaveOptions
            {
                // Ensure the format is set to Markdown (optional, default is Markdown).
                SaveFormat = SaveFormat.Markdown,

                // Export tables as raw HTML to preserve complex structures.
                ExportAsHtml = MarkdownExportAsHtml.Tables,

                // Set image resolution to 300 DPI for higher quality images.
                ImageResolution = 300,

                // Do not embed the Aspose.Words generator name in the output.
                ExportGeneratorName = false,

                // Export list items using plain text (instead of Markdown syntax) as an example.
                ListExportMode = MarkdownListExportMode.PlainText,

                // Export OfficeMath objects as LaTeX.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex
            };

            // Save the document as Markdown using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for a console application.
        /// </summary>
        static void Main(string[] args)
        {
            // Expect two arguments: input DOCM path and output Markdown path.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: DocmToMarkdownConverter <input.docm> <output.md>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                DocmToMarkdownConverter.Convert(inputPath, outputPath);
                Console.WriteLine($"Conversion succeeded: '{outputPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
