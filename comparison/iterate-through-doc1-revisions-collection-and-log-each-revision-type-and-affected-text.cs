using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Compare the documents to generate revisions.
        original.Compare(revised, "Author", DateTime.Now);

        // Save the compared document as an artifact.
        original.Save("Compared.docx");

        // Iterate through the revisions and log each revision type and its affected text.
        foreach (Revision rev in original.Revisions)
        {
            // The revision's parent node contains the actual text that was changed.
            string affectedText = rev.ParentNode?.GetText() ?? string.Empty;
            Console.WriteLine($"{rev.RevisionType}: {affectedText}");
        }
    }
}
