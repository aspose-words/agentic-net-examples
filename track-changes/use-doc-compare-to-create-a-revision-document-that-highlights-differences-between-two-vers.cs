using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It contains several lines of text.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Save the original for reference (optional).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Create the edited version of the document.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited document."); // Modified line.
        builder.Writeln("It contains several lines of text."); // Same line.
        builder.Writeln("The quick brown fox jumps over the lazy cat."); // Modified word.

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || edited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the edited document to the original.
        // The original document will now contain revisions that represent the differences.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Output revision details to the console.
        Console.WriteLine("Revisions detected:");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that highlights the differences (revisions are visible in Word).
        string revisionPath = Path.Combine(Directory.GetCurrentDirectory(), "Original_With_Revisions.docx");
        original.Save(revisionPath);

        // Accept all revisions to transform the original into the edited version.
        original.Revisions.AcceptAll();

        // Save the final accepted document.
        string finalPath = Path.Combine(Directory.GetCurrentDirectory(), "Final_Accepted.docx");
        original.Save(finalPath);
    }
}
