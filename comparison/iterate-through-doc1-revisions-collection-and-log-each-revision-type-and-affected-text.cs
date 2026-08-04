using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class RevisionLogger
{
    public static void Main()
    {
        // Create the original document with some content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Hello world!");
        builder1.Writeln("This is the original document.");

        // Create the revised document with modifications.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Hello brave new world!"); // changed line
        builder2.Writeln("This is the original document."); // unchanged line
        builder2.Writeln("An extra line added."); // insertion

        // Perform comparison – revisions will be added to doc1.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (doc1.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Log each revision's type and the affected text.
        foreach (Revision rev in doc1.Revisions)
        {
            string affectedText = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "[No node]";
            Console.WriteLine($"Revision type: {rev.RevisionType}, affected text: \"{affectedText}\"");
        }

        // Save the compared document (optional artifact).
        doc1.Save("Compared.docx");
    }
}
