using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || revised.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the result with revisions included.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedResult.docx");
        original.Save(outputPath);
    }
}
