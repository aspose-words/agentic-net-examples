using System;
using Aspose.Words;

namespace DocumentProcessing
{
    public class AppendixMerger
    {
        /// <summary>
        /// Merges the appendix document into the template when the flag is true.
        /// </summary>
        /// <param name="templatePath">Path to the main template DOCX file.</param>
        /// <param name="appendixPath">Path to the appendix DOCX file.</param>
        /// <param name="outputPath">Path where the merged document will be saved.</param>
        /// <param name="includeAppendix">Flag indicating whether the appendix should be appended.</param>
        public static void MergeIfNeeded(string templatePath, string appendixPath, string outputPath, bool includeAppendix)
        {
            // Load the main template document.
            Document template = new Document(templatePath);

            if (includeAppendix)
            {
                // Load the appendix document.
                Document appendix = new Document(appendixPath);

                // Append the appendix to the end of the template preserving its formatting.
                template.AppendDocument(appendix, ImportFormatMode.KeepSourceFormatting);
            }

            // Save the resulting document.
            template.Save(outputPath);
        }
    }

    // Entry point required for a console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Example usage – replace paths with real files when testing.
            string templatePath = "Template.docx";
            string appendixPath = "Appendix.docx";
            string outputPath = "MergedResult.docx";
            bool includeAppendix = true; // or false depending on your scenario

            AppendixMerger.MergeIfNeeded(templatePath, appendixPath, outputPath, includeAppendix);
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
