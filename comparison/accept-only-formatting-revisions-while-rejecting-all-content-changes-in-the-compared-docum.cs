using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with bold formatting on "Hello"
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Font.Bold = true;
        builderOriginal.Write("Hello");
        builderOriginal.Font.Bold = false;
        builderOriginal.Writeln(" world.");

        // Save the original for reference (optional)
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Create the revised document with italic formatting on "Hello" and an extra paragraph
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Font.Italic = true;
        builderRevised.Write("Hello");
        builderRevised.Font.Italic = false;
        builderRevised.Writeln(" world.");
        builderRevised.Writeln("Additional content added.");

        // Save the revised for reference (optional)
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        revised.Save(revisedPath);

        // Compare the documents; revisions are stored in the original document
        original.Compare(revised, "Comparer", DateTime.Now);

        Console.WriteLine($"Revisions after comparison: {original.Revisions.Count}");

        // Accept only formatting revisions, reject all others (insertions, deletions, moves)
        for (int i = original.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = original.Revisions[i];
            bool isFormatting = rev.RevisionType.ToString() == "Formatting";

            if (isFormatting)
                rev.Accept();
            else
                rev.Reject();
        }

        Console.WriteLine($"Revisions after processing: {original.Revisions.Count}");

        // Save the resulting document
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        original.Save(resultPath);
    }
}
