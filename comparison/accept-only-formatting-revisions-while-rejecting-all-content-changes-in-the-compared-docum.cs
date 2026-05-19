using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with bold formatting.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Font.Bold = true;
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a content insertion and different formatting.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        // Insert an extra word to create a content change.
        builderRevised.Writeln("Hello brave world.");
        // Ensure the text is not bold to create a formatting change.
        builderRevised.Font.Bold = false;

        // Compare the documents. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Accept only formatting revisions, reject all other types.
        // Copy the revisions to a list to avoid modifying the collection while iterating.
        List<Revision> revisions = new List<Revision>(original.Revisions);
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        original.Save(outputPath);
    }
}
