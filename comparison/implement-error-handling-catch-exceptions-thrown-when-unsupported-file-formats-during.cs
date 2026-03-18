using System;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsComparisonExample
{
    public class DocumentComparer
    {
        /// <summary>
        /// Loads two documents, handles unsupported format errors, compares them and saves the result.
        /// </summary>
        /// <param name="originalPath">Path to the original document.</param>
        /// <param name="revisedPath">Path to the document to compare against.</param>
        /// <param name="outputPath">Path where the comparison result will be saved.</param>
        public void CompareDocuments(string originalPath, string revisedPath, string outputPath)
        {
            Document? docOriginal = null;
            Document? docRevised = null;

            // Attempt to load the original document.
            try
            {
                LoadOptions loadOptions = new LoadOptions();
                docOriginal = new Document(originalPath, loadOptions);
            }
            catch (UnsupportedFileFormatException ex)
            {
                Console.WriteLine($"Unsupported file format for original document: {ex.Message}");
                return; // Abort comparison because the source cannot be loaded.
            }
            catch (FileCorruptedException ex)
            {
                Console.WriteLine($"Original document is corrupted: {ex.Message}");
                return;
            }

            // Attempt to load the revised document.
            try
            {
                LoadOptions loadOptions = new LoadOptions();
                docRevised = new Document(revisedPath, loadOptions);
            }
            catch (UnsupportedFileFormatException ex)
            {
                Console.WriteLine($"Unsupported file format for revised document: {ex.Message}");
                return; // Abort comparison because the target cannot be loaded.
            }
            catch (FileCorruptedException ex)
            {
                Console.WriteLine($"Revised document is corrupted: {ex.Message}");
                return;
            }

            // Ensure both documents are loaded before proceeding.
            if (docOriginal == null || docRevised == null)
            {
                Console.WriteLine("One or both documents could not be loaded.");
                return;
            }

            // Perform the comparison. The original document will receive revision marks.
            docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

            // Save the comparison result.
            try
            {
                docOriginal.Save(outputPath);
                Console.WriteLine($"Comparison completed and saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save comparison result: {ex.Message}");
            }
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point required by the .NET runtime. Demonstrates usage of DocumentComparer.
        /// </summary>
        /// <param name="args">Command‑line arguments: originalPath revisedPath outputPath</param>
        public static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: DocumentComparer <originalPath> <revisedPath> <outputPath>");
                return;
            }

            string originalPath = args[0];
            string revisedPath = args[1];
            string outputPath = args[2];

            var comparer = new DocumentComparer();
            comparer.CompareDocuments(originalPath, revisedPath, outputPath);
        }
    }
}
