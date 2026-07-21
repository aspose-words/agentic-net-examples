using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonShowDeletedContent
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Paragraph 1.");
        builderOriginal.Writeln("Paragraph to be deleted.");

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);

        // Delete the second paragraph in the revised document without tracking revisions.
        // This creates a difference that the comparison will treat as a deletion.
        revised.FirstSection.Body.Paragraphs[1].Remove();

        // Perform the comparison. The original document will receive the revisions.
        original.Compare(revised, "Author", DateTime.Now);

        // Verify that at least one revision (the deletion) exists.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Configure revision display options to retain deleted text in the output.
        original.LayoutOptions.RevisionOptions.ShowOriginalRevision = true;

        // Save the comparison result.
        original.Save("ComparisonResult.docx");
    }
}
