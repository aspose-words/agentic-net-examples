using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some formatted text.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is a sample paragraph.");
        builder.Font.Bold = true; // Apply bold formatting.
        builder.Writeln("Bold text line.");

        // Create the revised document that changes formatting and adds content.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("This is a sample paragraph."); // Same text.
        builder2.Font.Italic = true; // Change formatting (italic instead of bold).
        builder2.Writeln("Bold text line."); // Same text, different formatting.
        builder2.Writeln("Additional content line."); // New content.

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Accept only formatting revisions; reject all other types (insertions, deletions, etc.).
        var revisions = original.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();
            else
                rev.Reject();
        }

        // Verify that no revisions remain after processing.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Unexpected revisions remain after processing.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        original.Save(outputPath);
    }
}
