using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This line is unchanged.");

        // Create the revised document that contains differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world."); // Modified line.
        builder2.Writeln("This line is unchanged."); // Same line.
        builder2.Writeln("An extra line added in the revised version."); // New line.

        // Compare the documents, providing author name and comparison date.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document as DOC format while preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.doc");
        original.Save(outputPath, SaveFormat.Doc);
    }
}
