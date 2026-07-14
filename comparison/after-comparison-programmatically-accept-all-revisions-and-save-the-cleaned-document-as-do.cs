using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This line will stay unchanged.");

        // Create the revised document that contains a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!"); // changed text
        builderRevised.Writeln("This line will stay unchanged.");

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that the comparison produced revisions.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Accept all revisions, turning the original into the revised version.
        original.AcceptAllRevisions();

        // Ensure that all revisions have been accepted.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should have been accepted.");

        // Save the cleaned document.
        const string outputPath = "CleanedDocument.docx";
        original.Save(outputPath);
    }
}
