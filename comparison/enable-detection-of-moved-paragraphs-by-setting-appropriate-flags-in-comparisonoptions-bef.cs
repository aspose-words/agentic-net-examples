using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DetectMovedParagraphs
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2.");
        builder.Writeln("Paragraph 3.");

        // Clone the original to create the revised document.
        Document revised = (Document)original.Clone(true);

        // Move the second paragraph to the end of the document to create a move revision.
        // InsertAfter moves the node from its current position.
        Paragraph secondParagraph = revised.FirstSection.Body.Paragraphs[1];
        Paragraph thirdParagraph = revised.FirstSection.Body.Paragraphs[2];
        revised.FirstSection.Body.InsertAfter(secondParagraph, thirdParagraph);

        // Enable move detection in comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true
        };

        // Perform the comparison. The original document will receive the revisions.
        original.Compare(revised, "DemoAuthor", DateTime.Now, compareOptions);

        // Save the comparison result.
        string outputPath = "Compared.docx";
        original.Save(outputPath);

        // Report moved paragraph revisions.
        Console.WriteLine("Moved paragraph revisions detected:");
        ParagraphCollection paragraphs = original.FirstSection.Body.Paragraphs;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            if (para.IsMoveFromRevision)
                Console.WriteLine($"Paragraph {i + 1} is a Move-From revision: \"{para.GetText().Trim()}\"");
            if (para.IsMoveToRevision)
                Console.WriteLine($"Paragraph {i + 1} is a Move-To revision: \"{para.GetText().Trim()}\"");
        }

        // Additionally, list all revisions of type Moving.
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.Moving)
                Console.WriteLine($"Revision (Moving) on node type {rev.ParentNode.NodeType}: \"{rev.ParentNode.GetText().Trim()}\"");
        }
    }
}
