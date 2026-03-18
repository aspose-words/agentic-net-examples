using System;
using System.Collections.Generic;
using Aspose.Words;

namespace DocumentCombinerExample
{
    public static class DocumentCombiner
    {
        /// <summary>
        /// Combines multiple DOCX files into a single PDF document while preserving the original formatting of each source file.
        /// </summary>
        /// <param name="sourceDocxPaths">Enumerable collection of full file paths to the source DOCX files.</param>
        /// <param name="outputPdfPath">Full file path where the resulting PDF will be saved.</param>
        public static void CombineDocsToPdf(IEnumerable<string> sourceDocxPaths, string outputPdfPath)
        {
            // Validate input.
            if (sourceDocxPaths == null) throw new ArgumentNullException(nameof(sourceDocxPaths));
            if (string.IsNullOrEmpty(outputPdfPath)) throw new ArgumentException("Output path must be provided.", nameof(outputPdfPath));

            Document? combinedDocument = null;
            bool firstDocument = true;

            // Iterate through each DOCX file, load it, and append it to the combined document.
            foreach (string docxPath in sourceDocxPaths)
            {
                // Load the source document using the Document(string) constructor.
                Document srcDoc = new Document(docxPath);

                if (firstDocument)
                {
                    // Use the first document as the base for the combined document.
                    combinedDocument = srcDoc;
                    firstDocument = false;
                }
                else
                {
                    // Append subsequent documents while keeping their original formatting.
                    combinedDocument.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
                }
            }

            // If no documents were provided, throw an exception.
            if (combinedDocument == null)
                throw new InvalidOperationException("No source documents were provided.");

            // Save the combined document as PDF. The Save(string) overload determines the format from the file extension.
            combinedDocument!.Save(outputPdfPath);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point for the console application.
        /// Example usage: dotnet run -- "doc1.docx" "doc2.docx" "output.pdf"
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <input1.docx> [<input2.docx> ...] <output.pdf>");
                return;
            }

            // The last argument is the output PDF path; all preceding arguments are source DOCX files.
            string outputPdf = args[^1];
            var sourceDocs = args[..^1];

            try
            {
                DocumentCombiner.CombineDocsToPdf(sourceDocs, outputPdf);
                Console.WriteLine($"Successfully combined {sourceDocs.Length} document(s) into '{outputPdf}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
