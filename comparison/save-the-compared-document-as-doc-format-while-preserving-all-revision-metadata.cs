using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a paragraph.");

        // Create the edited document that has intentional differences.
        Document edited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(edited);
        builderEdited.Writeln("This is the edited document."); // Modified line.
        builderEdited.Writeln("It contains a paragraph.");      // Same line.
        builderEdited.Writeln("Additional line added.");       // New line.

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Perform the comparison, providing author name and timestamp.
            original.Compare(edited, "Comparer", DateTime.Now);
        }

        // Verify that revisions were generated.
        Console.WriteLine($"Revisions count after comparison: {original.Revisions.Count}");

        // Save the compared document (which now holds revision metadata) as DOC format.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.doc");
        original.Save(outputPath, SaveFormat.Doc);
    }
}
