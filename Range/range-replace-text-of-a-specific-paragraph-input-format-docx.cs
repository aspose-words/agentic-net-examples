using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace ParagraphReplaceExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example parameters – adjust paths as needed.
            string inputPath = @"C:\Docs\Source.docx";
            string outputPath = @"C:\Docs\Result.docx";
            int paragraphIndex = 2; // third paragraph (zero‑based)
            string oldText = "_FullName_";
            string newText = "John Doe";

            try
            {
                ReplaceParagraphText(inputPath, outputPath, paragraphIndex, oldText, newText);
                Console.WriteLine("Document saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// Replaces a text pattern inside a specific paragraph of a DOCX file.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOCX document.</param>
        /// <param name="outputPath">Full path where the modified document will be saved.</param>
        /// <param name="paragraphIndex">Zero‑based index of the paragraph to modify.</param>
        /// <param name="oldText">Text to be replaced.</param>
        /// <param name="newText">Replacement text.</param>
        public static void ReplaceParagraphText(string inputPath, string outputPath, int paragraphIndex, string oldText, string newText)
        {
            // Load the existing document.
            Document doc = new Document(inputPath);

            // Validate paragraph index.
            if (paragraphIndex < 0 || paragraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                throw new ArgumentOutOfRangeException(nameof(paragraphIndex), "Paragraph index is out of range.");

            // Get the target paragraph.
            Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[paragraphIndex];

            // Replace text only within this paragraph.
            int replacementsMade = targetParagraph.Range.Replace(oldText, newText, new FindReplaceOptions(FindReplaceDirection.Forward));

            if (replacementsMade == 0)
                Console.WriteLine("No occurrences of the pattern were found in the specified paragraph.");

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
