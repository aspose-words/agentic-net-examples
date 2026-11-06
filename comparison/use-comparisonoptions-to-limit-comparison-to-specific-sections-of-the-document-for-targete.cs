using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document with a header and body.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);

        // Add header content.
        builderOriginal.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builderOriginal.Writeln("Header original");

        // Add body content.
        builderOriginal.MoveToDocumentEnd();
        builderOriginal.Writeln("Body original");

        // Create the revised document with different header and body text.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);

        // Add revised header content.
        builderRevised.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builderRevised.Writeln("Header revised");

        // Add revised body content.
        builderRevised.MoveToDocumentEnd();
        builderRevised.Writeln("Body revised");

        // Configure compare options to ignore header/footer changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreHeadersAndFooters = true
        };

        // Perform the comparison. Header changes will be ignored; only body changes generate revisions.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Output revision information to the console.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Total revisions detected (excluding ignored sections): {revisionCount}");

        // List each revision type and its text.
        foreach (Revision revision in original.Revisions)
        {
            string revType = revision.RevisionType.ToString();
            // The revised text is stored in the node that the revision refers to.
            string revText = revision.ParentNode?.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"- {revType}: \"{revText}\"");
        }

        // Save the compared document showing the revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
