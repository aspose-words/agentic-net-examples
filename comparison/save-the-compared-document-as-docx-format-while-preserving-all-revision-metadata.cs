using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        var original = new Document();
        var builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This line will be deleted.");

        // Create the revised document that has intentional differences.
        var revised = new Document();
        var builderRevised = new DocumentBuilder(revised);
        // Modify the first line.
        builderRevised.Writeln("Hello world! Modified.");
        // Omit the line that existed in the original (deletion).
        // Add a new line (insertion).
        builderRevised.Writeln("This line is new.");

        // Ensure both documents are free of revisions before comparison.
        if (original.Revisions.Count != 0 || revised.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Perform the comparison. Revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Save the compared document preserving all revision metadata.
        const string outputFile = "ComparedWithRevisions.docx";
        original.Save(outputFile);
    }
}
