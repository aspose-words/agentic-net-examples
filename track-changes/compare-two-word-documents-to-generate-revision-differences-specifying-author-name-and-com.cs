using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a sample paragraph.");
        original.Save("Original.docx");

        // Create the edited document with some changes.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("Hello world!"); // unchanged line
        builder.Writeln("This is a modified paragraph."); // changed text
        builder.Writeln("Additional line."); // new line
        edited.Save("Edited.docx");

        // Load the documents for comparison.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Author name and comparison date for the revisions.
        string author = "Comparer";
        DateTime compareDate = DateTime.Now;

        // Ensure both documents have no existing revisions.
        if (docOriginal.HasRevisions || docEdited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents; revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, author, compareDate);

        // Output revision details.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision: Type={rev.RevisionType}, Author={rev.Author}, Date={rev.DateTime}");
            Console.WriteLine($"   Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that now contains the revision differences.
        docOriginal.Save("Compared.docx");
    }
}
