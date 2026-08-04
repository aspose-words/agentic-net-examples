using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DetectMovedParagraphs
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Clone the original to create the revised document.
        Document revised = (Document)original.Clone(true);

        // Move the second paragraph before the first paragraph in the revised document.
        // This will generate a move revision when compared with CompareMoves enabled.
        Node firstParagraph = revised.FirstSection.Body.Paragraphs[0];
        Node secondParagraph = revised.FirstSection.Body.Paragraphs[1];
        revised.FirstSection.Body.InsertBefore(secondParagraph, firstParagraph);

        // Set comparison options to detect moved paragraphs.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true,               // Enable move detection.
            Target = ComparisonTargetType.New // Use the revised document as the target.
        };

        // Perform the comparison.
        original.Compare(revised, "DemoAuthor", DateTime.Now, compareOptions);

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MovedParagraphs.docx");
        original.Save(outputPath);

        // Count and display move revisions.
        int moveRevisions = 0;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.Moving)
                moveRevisions++;
        }

        Console.WriteLine($"Total move revisions detected: {moveRevisions}");

        // Identify which paragraphs are part of a move revision.
        ParagraphCollection paragraphs = original.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            if (para.IsMoveFromRevision)
                Console.WriteLine($"Paragraph {i + 1} is a 'Move From' revision: \"{para.GetText().Trim()}\"");
            else if (para.IsMoveToRevision)
                Console.WriteLine($"Paragraph {i + 1} is a 'Move To' revision: \"{para.GetText().Trim()}\"");
        }
    }
}
