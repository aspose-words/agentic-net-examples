using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This is the original document.");
        builder.Writeln("It has two paragraphs.");

        // Create the edited document that contains differences.
        Document docEdited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(docEdited);
        builderEdited.Writeln("This is the edited document."); // changed first line
        builderEdited.Writeln("It has three paragraphs.");   // changed second line
        builderEdited.Writeln("This is an added third paragraph."); // new paragraph

        // Compare the original document with the edited one.
        // The original document will now contain revisions representing the differences.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Output the number of revisions detected after comparison.
        Console.WriteLine($"Revisions after compare: {docOriginal.Revisions.Count}");

        // Accept all revisions to produce a cleaned document without tracked changes.
        docOriginal.Revisions.AcceptAll();

        // Verify that all revisions have been accepted.
        Console.WriteLine($"Revisions after accept: {docOriginal.Revisions.Count}");

        // Save the cleaned document as DOCX.
        docOriginal.Save("CleanedDocument.docx");
    }
}
