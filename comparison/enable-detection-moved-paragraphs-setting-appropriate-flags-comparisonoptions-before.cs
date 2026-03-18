using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DetectMovedParagraphs
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options to include move revisions.
        CompareOptions compareOptions = new CompareOptions
        {
            // Enable detection of moved paragraphs.
            CompareMoves = true,
            // Keep other options at their defaults (no ignoring).
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

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Iterate through paragraphs to find move revisions.
        ParagraphCollection paragraphs = docOriginal.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];

            if (para.IsMoveFromRevision)
            {
                Console.WriteLine($"Paragraph {i} is a 'move from' revision:");
                Console.WriteLine(para.GetText());
            }
            else if (para.IsMoveToRevision)
            {
                Console.WriteLine($"Paragraph {i} is a 'move to' revision:");
                Console.WriteLine(para.GetText());
            }
        }

        // Save the document with revisions.
        docOriginal.Save("ComparisonResult.docx");
    }
}
