using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class CompareAndSaveDoc
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world!");
        builder1.Writeln("This line will stay the same.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello brave new world!"); // Modified line.
        builder2.Writeln("This line will stay the same."); // Unchanged line.
        builder2.Writeln("Additional line added."); // New line.

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "JD", DateTime.Now);

        // Ensure that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were generated after comparison.");
        }

        // Save the compared document as a binary DOC file, preserving all revision metadata.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.doc");
        original.Save(outputPath, SaveFormat.Doc);

        // Reload the saved document to verify that revisions are still present.
        Document loaded = new Document(outputPath);
        if (loaded.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Revisions were lost after saving as DOC.");
        }

        // Indicate successful completion.
        Console.WriteLine("Comparison saved successfully with revisions.");
    }
}
