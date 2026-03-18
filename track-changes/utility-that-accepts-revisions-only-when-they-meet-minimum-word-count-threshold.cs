using System;
using Aspose.Words;

namespace RevisionUtility
{
    // Custom criteria that matches revisions whose text contains at least a given number of words.
    public class WordCountRevisionCriteria : IRevisionCriteria
    {
        private readonly int _minWordCount;

        public WordCountRevisionCriteria(int minWordCount)
        {
            _minWordCount = minWordCount;
        }

        // Returns true if the revision's text has >= _minWordCount words.
        public bool IsMatch(Revision revision)
        {
            // Get the text of the node that the revision is attached to.
            string text = revision.ParentNode?.GetText() ?? string.Empty;

            // Simple word count: split on whitespace and count non‑empty entries.
            int wordCount = 0;
            foreach (var part in text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                wordCount++;

            return wordCount >= _minWordCount;
        }
    }

    public static class RevisionProcessor
    {
        // Accepts only revisions that meet the minimum word count; all others are rejected.
        public static void AcceptRevisionsByWordCount(string inputPath, string outputPath, int minWordCount)
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Accept revisions that satisfy the word‑count criteria.
            doc.Revisions.Accept(new WordCountRevisionCriteria(minWordCount));

            // Reject any remaining revisions (those that did not meet the criteria).
            doc.Revisions.RejectAll();

            // Save the processed document.
            doc.Save(outputPath);
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            string sourceFile = @"C:\Docs\Input.docx";
            string resultFile = @"C:\Docs\Output.docx";
            int minimumWords = 5; // Minimum word count threshold.

            RevisionProcessor.AcceptRevisionsByWordCount(sourceFile, resultFile, minimumWords);
        }
    }
}
