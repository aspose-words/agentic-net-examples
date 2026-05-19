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

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words world!");

        // Compare the documents – this adds revisions to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions – this removes the revision marks.
        original.Revisions.AcceptAll();

        // After acceptance, the revisions collection should be empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not cleared after acceptance.");

        // Save the resulting document.
        original.Save("final.docx");
    }
}
