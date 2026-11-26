using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a footnote.
        builder.Writeln("This is a paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Add a paragraph with an endnote.
        builder.Writeln("This is a paragraph with an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote text.");

        // Save the original document.
        doc.Save("Original.docx");

        // Remove all footnotes while preserving endnotes.
        var footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnoteNodes.Cast<Footnote>().ToList())
        {
            if (footnote.FootnoteType == FootnoteType.Footnote)
                footnote.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");

        // Output counts for verification.
        int remainingFootnotes = doc.GetChildNodes(NodeType.Footnote, true)
                                    .Cast<Footnote>()
                                    .Count(fn => fn.FootnoteType == FootnoteType.Footnote);
        int remainingEndnotes = doc.GetChildNodes(NodeType.Footnote, true)
                                   .Cast<Footnote>()
                                   .Count(fn => fn.FootnoteType == FootnoteType.Endnote);

        Console.WriteLine($"Footnotes remaining: {remainingFootnotes}");
        Console.WriteLine($"Endnotes remaining: {remainingEndnotes}");
    }
}
