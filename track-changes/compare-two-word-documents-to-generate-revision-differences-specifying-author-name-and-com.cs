using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It has two paragraphs.");
        // Save the original for reference (optional).
        original.Save("Original.docx");

        // Create the edited document with some changes.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited document."); // changed first line
        builder.Writeln("It has three paragraphs now."); // changed second line
        builder.Writeln("Additional paragraph added.");   // new paragraph
        // Save the edited for reference (optional).
        edited.Save("Edited.docx");

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || edited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The differences will appear as revisions in the original document.
        string authorName = "John Doe";
        DateTime comparisonDate = DateTime.Now;
        original.Compare(edited, authorName, comparisonDate);

        // Output revision details to the console.
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Author: {rev.Author}");
            Console.WriteLine($"Date: {rev.DateTime}");
            Console.WriteLine($"Type: {rev.RevisionType}");
            Console.WriteLine($"Text: \"{rev.ParentNode.GetText().Trim()}\"");
            Console.WriteLine(new string('-', 40));
        }

        // Save the document that now contains the revision markup.
        original.Save("Compared.docx");
    }
}
