using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This is a sample paragraph.");

        // Create the revised document that contains differences.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world."); // Modified line.
        builder2.Writeln("This is a sample paragraph."); // Unchanged line.
        builder2.Writeln("An added line."); // New line.

        // Perform comparison; revisions are stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the compared document (optional artifact).
        original.Save("Compared.docx");

        // Iterate through all revisions and log their type and affected text.
        foreach (Revision revision in original.Revisions)
        {
            // The affected text can be obtained from the node that the revision is attached to.
            string affectedText = revision.ParentNode?.GetText() ?? string.Empty;
            Console.WriteLine($"{revision.RevisionType}: {affectedText}");
        }
    }
}
