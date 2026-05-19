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
        builder1.Writeln("This is the original document.");

        // Create the revised document with different content.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello brave new world.");
        builder2.Writeln("This is the edited document.");

        // Compare the documents; revisions are added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Save the compared document (optional artifact).
        original.Save("Compared.docx");

        // Iterate through the revisions and log their type and affected text.
        foreach (Revision rev in original.Revisions)
        {
            string type = rev.RevisionType.ToString();
            string text = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"Revision type: {type}, affected text: \"{text}\"");
        }
    }
}
