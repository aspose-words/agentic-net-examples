using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class RevisionComparisonExample
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It contains a few lines of text.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Create the edited document with some changes.
        Document edited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(edited);
        builderEdited.Writeln("This is the edited document."); // changed line
        builderEdited.Writeln("It contains a few lines of text."); // unchanged
        builderEdited.Writeln("The quick brown fox jumps over the lazy cat."); // changed word

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will now contain revisions that represent the differences.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Output revision details to the console.
        Console.WriteLine("Revisions found after comparison:");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}");
            Console.WriteLine($"  Changed text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the original document (which now includes revision markup) to a file.
        string outputPath = "OriginalWithRevisions.docx";
        original.Save(outputPath);
        Console.WriteLine($"Revision document saved to: {outputPath}");
    }
}
