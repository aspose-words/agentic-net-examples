using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains a few sentences.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original paragraph."); // same line
        builderRevised.Writeln("It contains several modified sentences."); // changed text
        builderRevised.Writeln("An additional line is added."); // new line

        // Perform the comparison. The original document will receive revision objects.
        string author = "AsposeUser";
        DateTime compareDate = DateTime.Now;
        original.Compare(revised, author, compareDate);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the compared document (with revisions) as DOCX.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(outputPath, SaveFormat.Docx);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Comparison complete. Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
