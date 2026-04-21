using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class CompareDocumentsExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It contains several lines of text.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Create the edited document.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited document."); // changed line
        builder.Writeln("It contains several lines of text."); // unchanged
        builder.Writeln("The quick brown fox jumps over the lazy cat."); // changed word

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || edited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will receive revisions that represent the differences.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Output revision information to the console.
        Console.WriteLine($"Number of revisions detected: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Author: {rev.Author}");
            Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the original document with revisions highlighted.
        string resultPath = Path.Combine(outputDir, "OriginalWithRevisions.docx");
        original.Save(resultPath);
        Console.WriteLine($"Revised document saved to: {resultPath}");
    }
}
