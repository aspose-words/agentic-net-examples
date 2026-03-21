using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

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
            // Create a temporary input document with tracked revisions.
            string sourceFile = Path.Combine(Path.GetTempPath(), $"Input_{Guid.NewGuid()}.docx");
            string resultFile = Path.Combine(Path.GetTempPath(), $"Output_{Guid.NewGuid()}.docx");
            int minimumWords = 5; // Minimum word count threshold.

            // Build a document with revisions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start tracking revisions.
            doc.StartTrackRevisions("Author");

            // This revision has enough words (>=5).
            builder.Writeln("This is a revision that contains enough words.");

            // This revision does NOT meet the threshold.
            builder.Writeln("Short rev.");

            // Stop tracking.
            doc.StopTrackRevisions();

            // Save the input document.
            doc.Save(sourceFile);

            // Process the document using the utility.
            RevisionProcessor.AcceptRevisionsByWordCount(sourceFile, resultFile, minimumWords);

            // Output paths for verification (optional).
            Console.WriteLine($"Input document:  {sourceFile}");
            Console.WriteLine($"Processed document: {resultFile}");
        }
    }
}
