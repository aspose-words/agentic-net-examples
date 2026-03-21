using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsComparisonValidator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the original document using a DocumentBuilder that does not require a file path.
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Hello world!");
            builder.Writeln("This is the original document.");
            builder.Writeln("It has three paragraphs.");
            Document originalDoc = builder.Document;

            // Create the edited document based on the original.
            Document editedDoc = (Document)originalDoc.Clone();

            // Insert a new paragraph (insertion) at the end of the edited document.
            Paragraph newPara = new Paragraph(editedDoc);
            Run run = new Run(editedDoc, "This paragraph was added in the edited version.");
            newPara.AppendChild(run);
            editedDoc.FirstSection.Body.AppendChild(newPara);

            // Delete a paragraph (deletion) by removing the first paragraph from the original.
            originalDoc.FirstSection.Body.Paragraphs[0].Remove();

            // Perform validation/comparison.
            ValidateComparison(originalDoc, editedDoc);
        }

        /// <summary>
        /// Compares two documents and prints the resulting revisions.
        /// </summary>
        /// <param name="original">The original document (will receive revisions).</param>
        /// <param name="edited">The edited document to compare against.</param>
        static void ValidateComparison(Document original, Document edited)
        {
            // Ensure both documents are revision‑free before comparison.
            if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
                throw new InvalidOperationException("Both documents must not contain revisions before comparison.");

            // Set up comparison options (default values are sufficient for a full comparison).
            CompareOptions compareOptions = new CompareOptions
            {
                CompareMoves = false,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to the original document.
            original.Compare(edited, "Validator", DateTime.Now, compareOptions);

            // Output revision information.
            int revisionCount = original.Revisions.Count;
            Console.WriteLine("Comparison validation succeeded.");
            Console.WriteLine($"Total revisions: {revisionCount}");
            Console.WriteLine("Revision types:");
            foreach (Revision rev in original.Revisions)
                Console.WriteLine($" - {rev.RevisionType}");
        }
    }
}
