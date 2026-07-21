using System;
using Aspose.Words;

public class CompareDocumentsExample
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It has two paragraphs.");
        original.Save("Original.docx");

        // Create the edited document.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited document.");
        builder.Writeln("It has three paragraphs now.");
        builder.Writeln("Additional line added.");
        edited.Save("Edited.docx");

        // Compare the documents. The original document will receive revisions.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Output revision details.
        Console.WriteLine("Revisions found after comparison:");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so the original becomes identical to the edited version.
        original.Revisions.AcceptAll();

        // Verify that the texts now match.
        if (original.GetText() == edited.GetText())
        {
            Console.WriteLine("All revisions accepted. Documents are now identical.");
        }
        else
        {
            Console.WriteLine("Documents differ after accepting revisions.");
        }

        // Save the revised document.
        original.Save("Revised.docx");
    }
}
