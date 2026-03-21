using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This is a sample paragraph.");
        docOriginal.Save("Original.docx");

        // Create the edited document based on the original and apply a formatting change.
        Document docEdited = new Document("Original.docx");
        DocumentBuilder editBuilder = new DocumentBuilder(docEdited);
        editBuilder.MoveToDocumentStart();
        editBuilder.Font.Bold = true;
        editBuilder.Writeln("This is a sample paragraph."); // same text, different formatting
        docEdited.Save("Edited.docx");

        // Configure comparison options to track only formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = true,
            IgnoreComments = true,
            IgnoreTables = true,
            IgnoreFields = true,
            IgnoreFootnotes = true,
            IgnoreTextboxes = true,
            IgnoreHeadersAndFooters = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Accept only formatting revisions, reject all other content revisions.
        var revisions = docOriginal.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document with only formatting revisions applied.
        docOriginal.Save("Result.docx");
    }
}
