using System;
using Aspose.Words;

public class RevisionLogger
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This line will be changed.");
        builder1.Writeln("This line stays the same.");

        // Create the revised document that contains differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello world!");                     // Punctuation change.
        builder2.Writeln("This line has been modified.");    // Text change.
        builder2.Writeln("This line stays the same.");       // Unchanged.

        // Perform the comparison – revisions are added to the original document.
        original.Compare(revised, "John Doe", DateTime.Now);

        // Verify that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Log each revision's type and the affected text.
        foreach (Revision rev in original.Revisions)
        {
            // Some revision types (e.g., style changes) may not have a ParentNode.
            string affectedText = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no node>";
            Console.WriteLine($"Revision type: {rev.RevisionType}, affected text: \"{affectedText}\"");
        }

        // Save the compared document (optional artifact).
        original.Save("Compared.docx");
    }
}
