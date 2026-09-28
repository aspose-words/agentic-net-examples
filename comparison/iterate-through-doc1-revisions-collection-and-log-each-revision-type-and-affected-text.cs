using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the documents to generate revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Iterate through the revisions and log their type and affected text.
        foreach (Revision rev in original.Revisions)
        {
            // The affected text can be obtained from the revision's parent node.
            string text = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Text: \"{text}\"");
        }

        // Save the document that contains the revisions.
        original.Save("compared.docx");
    }
}
