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

    public static class Program
    {
        public static void Main(string[] args)
        {
            // Create simple placeholder documents so the example runs without external files.
            string templatePath = "Template.docx";
            string appendixPath = "Appendix.docx";
            string outputPath = "MergedResult.docx";
            bool includeAppendix = true; // Change to false to test without merging.

            // Generate a basic template document.
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("This is the main template document.");
            templateDoc.Save(templatePath);

            // Generate a basic appendix document.
            Document appendixDoc = new Document();
            DocumentBuilder appendixBuilder = new DocumentBuilder(appendixDoc);
            appendixBuilder.Writeln("This is the appendix document.");
            appendixDoc.Save(appendixPath);

            // Perform the merge based on the flag.
            AppendixMerger.MergeIfNeeded(templatePath, appendixPath, outputPath, includeAppendix);
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
