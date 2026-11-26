using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Create the first document (original) with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Save the original document to a local file.
        string originalPath = "original.docx";
        original.Save(originalPath, SaveFormat.Docx);

        // Create the second document (revised) with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Save the revised document to a local file.
        string revisedPath = "revised.docx";
        revised.Save(revisedPath, SaveFormat.Docx);

        // Load the saved documents from disk.
        Document docToCompare = new Document(originalPath);
        Document docToCompareAgainst = new Document(revisedPath);

        // Perform the comparison. Revisions will be added to docToCompare.
        docToCompare.Compare(docToCompareAgainst, "ComparisonAuthor", DateTime.Now);

        // Verify that revisions were created.
        if (docToCompare.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result, which includes the revisions.
        string resultPath = "comparison-result.docx";
        docToCompare.Save(resultPath, SaveFormat.Docx);
    }
}
