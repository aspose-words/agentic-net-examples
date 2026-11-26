using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This line will stay unchanged.");

        // Create the revised document that contains differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world."); // Modified line.
        builderRevised.Writeln("This line will stay unchanged."); // Same line.
        builderRevised.Writeln("An extra line added."); // New line.

        // Compare the two documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions, cleaning the document.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should have been accepted.");
        }

        // Save the cleaned document as DOCX.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "cleaned.docx");
        original.Save(outputPath, SaveFormat.Docx);
    }
}
