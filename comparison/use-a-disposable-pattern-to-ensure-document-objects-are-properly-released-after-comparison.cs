using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DocumentComparisonDemo
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparisonResult.docx");

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(original);
        originalBuilder.Writeln("This is the original paragraph.");
        original.Save(originalPath);

        // Create the revised document with a deliberate change.
        Document revised = new Document();
        DocumentBuilder revisedBuilder = new DocumentBuilder(revised);
        revisedBuilder.Writeln("This is the edited paragraph."); // Different text.
        revised.Save(revisedPath);

        // Load both documents, compare them.
        Document originalDoc = new Document(originalPath);
        Document revisedDoc = new Document(revisedPath);

        // Perform the comparison. The original document will receive revisions.
        originalDoc.Compare(revisedDoc, "DemoAuthor", DateTime.Now);

        // Verify that at least one revision was created.
        if (originalDoc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Output revision details to the console.
        foreach (Revision rev in originalDoc.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, author: {rev.Author}");
            Console.WriteLine($"Changed text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so the original document becomes identical to the revised one.
        originalDoc.AcceptAllRevisions();

        // After acceptance, there should be no remaining revisions.
        if (originalDoc.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should have been accepted.");

        // Save the final document.
        originalDoc.Save(resultPath);

        // Clean up temporary files (optional).
        File.Delete(originalPath);
        File.Delete(revisedPath);
    }
}
