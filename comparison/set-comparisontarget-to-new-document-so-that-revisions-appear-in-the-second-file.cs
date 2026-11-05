using System;
using Aspose.Words;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a line that will be changed.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original document.");
        // Changed line.
        builderRevised.Writeln("It contains a line that has been modified.");

        // Perform the comparison. Revisions will be added to the revised document
        // because we invoke Compare on the revised document.
        revised.Compare(original, "Comparer", DateTime.Now);

        // Verify that revisions exist in the revised document.
        if (revised.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions in the revised document, but none were found.");
        }

        // Save both documents.
        original.Save("Original.docx");
        revised.Save("Revised_With_Revisions.docx");
    }
}
