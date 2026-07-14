using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This line will stay the same.");

        // Create the revised document that contains intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world."); // Modified text.
        builderRevised.Writeln("This line will stay the same."); // Unchanged text.
        builderRevised.Writeln("Additional line added."); // New text.

        // Compare the two documents. Revisions will be added to the original document.
        original.Compare(revised, "DemoAuthor", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the document that now contains the revisions.
        original.Save("ComparedResult.docx");
    }
}
