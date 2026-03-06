using System;
using Aspose.Words;

namespace ReplaceParagraphTextExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths can be supplied via command‑line arguments, configuration, or hard‑coded for a quick demo.
            string inputPath = @"C:\Docs\Source.docx";
            string outputPath = @"C:\Docs\Result.docx";
            int paragraphIndex = 2;               // third paragraph (zero‑based)
            string oldText = "_FullName_";
            string newText = "John Doe";

            try
            {
                ParagraphTextReplacer.ReplaceParagraphText(inputPath, outputPath, paragraphIndex, oldText, newText);
                Console.WriteLine("Paragraph text replaced successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public static class ParagraphTextReplacer
    {
        /// <summary>
        /// Replaces occurrences of <paramref name="oldText"/> with <paramref name="newText"/> in the paragraph at
        /// <paramref name="paragraphIndex"/> of the document located at <paramref name="inputPath"/>.
        /// The modified document is saved to <paramref name="outputPath"/>.
        /// </summary>
        public static void ReplaceParagraphText(string inputPath, string outputPath, int paragraphIndex, string oldText, string newText)
        {
            // Load the existing document.
            Document doc = new Document(inputPath);

            // Validate the paragraph index.
            if (paragraphIndex < 0 || paragraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                throw new ArgumentOutOfRangeException(nameof(paragraphIndex), "Paragraph index is out of range.");

            // Get the target paragraph.
            Paragraph targetParagraph = doc.FirstSection.Body.Paragraphs[paragraphIndex];

            // Perform the replace operation inside the paragraph's range.
            int replacementsMade = targetParagraph.Range.Replace(oldText, newText);

            if (replacementsMade == 0)
                Console.WriteLine("No occurrences of the specified text were found in the selected paragraph.");

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
