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
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Paragraph 1 original.");
        builderOriginal.Writeln("Paragraph 2 original.");
        builderOriginal.Writeln("Paragraph 3 original.");

        // Create the revised document:
        // - Paragraph 2 text is changed.
        // - Paragraph 3 is removed.
        // - A new paragraph 4 is added.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Paragraph 1 original."); // unchanged
        builderRevised.Writeln("Paragraph 2 revised."); // modified
        builderRevised.Writeln("Paragraph 4 new.");     // added

        // Compare the documents. The original document will contain revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after comparison.");
        }

        // Summarize revisions by type for paragraph nodes.
        int addedParagraphs = 0;
        int deletedParagraphs = 0;
        int modifiedParagraphs = 0; // Format changes on paragraphs.

        foreach (Revision rev in original.Revisions)
        {
            // Consider only paragraph-level revisions.
            if (rev.ParentNode?.NodeType == NodeType.Paragraph)
            {
                switch (rev.RevisionType)
                {
                    case RevisionType.Insertion:
                        addedParagraphs++;
                        break;
                    case RevisionType.Deletion:
                        deletedParagraphs++;
                        break;
                    case RevisionType.FormatChange:
                        modifiedParagraphs++;
                        break;
                }
            }
        }

        // Build a simple text report.
        string report = $"Revision Summary:{Environment.NewLine}" +
                        $"Added paragraphs   : {addedParagraphs}{Environment.NewLine}" +
                        $"Deleted paragraphs : {deletedParagraphs}{Environment.NewLine}" +
                        $"Modified paragraphs: {modifiedParagraphs}{Environment.NewLine}";

        // Save the report to a local text file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionSummary.txt");
        File.WriteAllText(reportPath, report);

        // Save the compared document that contains the tracked changes.
        string comparedDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedDocPath);

        // Optionally write the report to the console (no user interaction required).
        Console.WriteLine(report);
    }
}
