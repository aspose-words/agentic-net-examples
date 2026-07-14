using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonShowDeletedContentExample
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("First paragraph.");
        builderOriginal.Writeln("Second paragraph to be deleted.");

        // Create the revised document that lacks the second paragraph.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("First paragraph.");

        // Set up compare options. No special flags are needed to keep deleted text visible;
        // deletion revisions are retained by default.
        CompareOptions compareOptions = new CompareOptions();

        // Perform the comparison. The original document will receive the revisions.
        original.Compare(revised, "Author", DateTime.Now, compareOptions);

        // Save the comparison result to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
