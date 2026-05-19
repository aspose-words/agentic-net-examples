using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains a line that will be changed.");

        // Create the revised document with modifications.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original paragraph.");
        // Change the second line.
        builderRevised.Writeln("It contains a line that has been edited.");

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document in the legacy DOC format, preserving revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.doc");
        original.Save(outputPath);
    }
}
