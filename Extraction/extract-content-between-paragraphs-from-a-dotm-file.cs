using System;
using System.IO;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        /// <summary>
        /// Extracts the combined text of paragraphs between the specified start and end indices
        /// (inclusive) from a DOTM (macro‑enabled template) file.
        /// </summary>
        /// <param name="dotmPath">Full path to the .dotm file.</param>
        /// <param name="startParagraphIndex">Zero‑based index of the first paragraph to include.</param>
        /// <param name="endParagraphIndex">Zero‑based index of the last paragraph to include.</param>
        /// <returns>Plain text containing the concatenated paragraph texts.</returns>
        static string ExtractBetweenParagraphs(string dotmPath, int startParagraphIndex, int endParagraphIndex)
        {
            // Load the DOTM file. The Document constructor automatically detects the format.
            Document doc = new Document(dotmPath);

            // Ensure the document has at least one section and body.
            if (doc.FirstSection == null || doc.FirstSection.Body == null)
                throw new InvalidOperationException("The document does not contain any body paragraphs.");

            // Get the collection of all paragraphs in the main story.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Validate indices.
            if (startParagraphIndex < 0 || endParagraphIndex < 0 ||
                startParagraphIndex >= paragraphs.Count || endParagraphIndex >= paragraphs.Count ||
                startParagraphIndex > endParagraphIndex)
                throw new ArgumentOutOfRangeException("Invalid paragraph index range.");

            // Accumulate the text of the selected paragraphs.
            StringWriter writer = new StringWriter();
            for (int i = startParagraphIndex; i <= endParagraphIndex; i++)
            {
                // GetText() returns the paragraph text including the trailing paragraph break.
                writer.Write(paragraphs[i].GetText());
            }

            return writer.ToString();
        }

        static void Main(string[] args)
        {
            // Example usage:
            string dotmFile = @"C:\Docs\Template.dotm";

            // Extract paragraphs 2 through 5 (zero‑based indices 1 to 4).
            string extractedText = ExtractBetweenParagraphs(dotmFile, 1, 4);

            // Output the result to console or save to a file.
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(extractedText);

            // Optionally, save the extracted text to a plain‑text file.
            File.WriteAllText(@"C:\Docs\ExtractedContent.txt", extractedText);
        }
    }
}
