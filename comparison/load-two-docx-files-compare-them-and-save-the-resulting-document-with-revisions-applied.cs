using System;
using Aspose.Words;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a single paragraph.");

        // Create the revised document with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document.");
        builderRevised.Writeln("It now contains two paragraphs.");
        builderRevised.Writeln("Additional line added for comparison.");

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the document that now contains the revisions.
        const string outputPath = "Compared.docx";
        original.Save(outputPath);
    }
}
