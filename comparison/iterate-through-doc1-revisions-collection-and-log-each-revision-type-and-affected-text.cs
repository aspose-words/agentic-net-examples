using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a slight change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the documents; revisions are added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Iterate through the revisions and log their type and affected text.
        foreach (Revision rev in original.Revisions)
        {
            string text = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"Revision type: {rev.RevisionType}, text: \"{text}\"");
        }

        // Save the resulting document with revisions.
        original.Save("Compared.docx");
    }
}
