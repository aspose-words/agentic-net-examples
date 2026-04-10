using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create the original document with two paragraphs.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Clone the original to create the edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // In the edited document, move the second paragraph before the first one.
        // This change will be detected as a move revision when CompareMoves is enabled.
        Paragraph secondParagraph = docEdited.FirstSection.Body.Paragraphs[1];
        Paragraph firstParagraph = docEdited.FirstSection.Body.Paragraphs[0];
        docEdited.FirstSection.Body.InsertBefore(secondParagraph, firstParagraph);

        // Set up comparison options to enable move detection.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true,                     // Enable move revisions.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New       // Use the edited document as the target.
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Save the document that now contains the revisions.
        string resultPath = Path.Combine(artifactsDir, "ComparisonResult.docx");
        docOriginal.Save(resultPath);

        // Output information about move revisions.
        int movingRevisions = docOriginal.Revisions.Count(r => r.RevisionType == RevisionType.Moving);
        Console.WriteLine($"Total moving revisions detected: {movingRevisions}");

        // Examine each paragraph to see if it is part of a move revision.
        ParagraphCollection paragraphs = docOriginal.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            if (para.IsMoveFromRevision)
                Console.WriteLine($"Paragraph {i} is a 'Move From' revision: \"{para.GetText().Trim()}\"");
            if (para.IsMoveToRevision)
                Console.WriteLine($"Paragraph {i} is a 'Move To' revision: \"{para.GetText().Trim()}\"");
        }
    }
}
