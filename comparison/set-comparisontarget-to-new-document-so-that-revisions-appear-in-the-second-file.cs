using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonTargetExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the edited document by cloning the original and changing its text.
        Document docEdited = (Document)docOriginal.Clone(true);
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "This is the edited paragraph.";

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.HasRevisions || docEdited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up compare options to use the new document as the comparison target.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be recorded in the second document (docEdited).
        docOriginal.Compare(docEdited, "AsposeUser", DateTime.Now, compareOptions);

        // Save both documents so we can inspect the result.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        string editedPath = Path.Combine(outputDir, "Edited_With_Revisions.docx");
        docOriginal.Save(originalPath);
        docEdited.Save(editedPath);

        // Verify that revisions appear in the edited document.
        int revisionCount = docEdited.Revisions.Count;
        Console.WriteLine($"Revisions in the edited document: {revisionCount}");

        // Optional: list revision details.
        foreach (Revision rev in docEdited.Revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }
    }
}
