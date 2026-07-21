using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This line will stay the same.");

        // Create the revised document with a change.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words!");
        builderRevised.Writeln("This line will stay the same.");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions.
        original.Revisions.AcceptAll();

        // After accepting, the revisions collection should be empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions should be cleared after acceptance.");

        // Save the resulting document.
        original.Save("Result.docx");
    }
}
