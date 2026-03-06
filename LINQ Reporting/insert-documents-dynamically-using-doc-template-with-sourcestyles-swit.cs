using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    public class DocumentInserter
    {
        /// <summary>
        /// Inserts a collection of source documents into a template document.
        /// If <paramref name="useSourceStyles"/> is true, the insertion uses
        /// KeepSourceFormatting mode with SmartStyleBehavior enabled to resolve
        /// style name clashes by converting source styles to direct formatting.
        /// </summary>
        /// <param name="templatePath">Path to the DOC template file.</param>
        /// <param name="sourcePaths">Paths to the source DOC files to be inserted.</param>
        /// <param name="useSourceStyles">When true, source formatting is preserved.</param>
        /// <param name="outputPath">Path where the resulting document will be saved.</param>
        public static void InsertDocuments(string templatePath, IEnumerable<string> sourcePaths, bool useSourceStyles, string outputPath)
        {
            // Load the template document.
            Document templateDoc = new Document(templatePath);

            // Create a DocumentBuilder for the template.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Iterate over each source document.
            foreach (string srcPath in sourcePaths)
            {
                // Load the source document.
                Document srcDoc = new Document(srcPath);

                // Prepare import options.
                ImportFormatOptions importOptions = new ImportFormatOptions();

                // If the caller wants to keep source styles, enable SmartStyleBehavior.
                // This expands source styles with the same name as destination styles
                // into direct formatting, avoiding style clashes.
                if (useSourceStyles)
                {
                    importOptions.SmartStyleBehavior = true;
                }

                // Move the cursor to the end of the template and insert a page break
                // to separate each inserted document.
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);

                // Insert the source document using KeepSourceFormatting mode.
                // The ImportFormatOptions are passed to control style handling.
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, importOptions);
            }

            // Save the combined document.
            templateDoc.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            string templateFile = @"C:\Docs\Template.docx";
            var sources = new List<string>
            {
                @"C:\Docs\Part1.docx",
                @"C:\Docs\Part2.docx",
                @"C:\Docs\Part3.docx"
            };
            bool preserveSourceStyles = true; // switch to control style handling
            string resultFile = @"C:\Docs\CombinedResult.docx";

            InsertDocuments(templateFile, sources, preserveSourceStyles, resultFile);
        }
    }
}
