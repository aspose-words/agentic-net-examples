using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1: Introduction.");
        builder.Writeln("Paragraph 2: Details.");
        builder.Writeln("Paragraph 3: Conclusion.");

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);

        // Move the second paragraph to the end to simulate a paragraph move.
        Paragraph paragraphToMove = revised.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"
        revised.FirstSection.Body.Paragraphs.RemoveAt(1);
        revised.FirstSection.Body.Paragraphs.Add(paragraphToMove);

        // Set comparison options to detect moved paragraphs.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true, // Enable move detection.
            // Other flags remain default (false) to keep other differences visible if any.
        };

        // Perform the comparison. The original document will contain the revisions.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "MovedParagraphsComparison.docx");
        original.Save(resultPath);

        // Inspect paragraphs for move revisions.
        ParagraphCollection paragraphs = original.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            if (para.IsMoveFromRevision)
            {
                Console.WriteLine($"Paragraph at index {i} is a moved-from revision: \"{para.GetText().Trim()}\"");
            }
            else if (para.IsMoveToRevision)
            {
                Console.WriteLine($"Paragraph at index {i} is a moved-to revision: \"{para.GetText().Trim()}\"");
            }
        }

        // Verify that at least one move revision was detected.
        bool hasMoveRevisions = original.Revisions.Any(r => r.RevisionType == RevisionType.Moving);
        if (!hasMoveRevisions)
            throw new InvalidOperationException("Expected at least one moving revision, but none were found.");
    }
}
