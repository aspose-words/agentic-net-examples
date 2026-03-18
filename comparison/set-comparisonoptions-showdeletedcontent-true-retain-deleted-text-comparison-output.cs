using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello world!");
        builder.Writeln("This paragraph will be deleted.");

        // Clone the original and delete the second paragraph while tracking revisions.
        Document docEdited = (Document)docOriginal.Clone(true);
        docEdited.StartTrackRevisions("John Doe", DateTime.Now);
        // Remove the paragraph that should appear as a deletion.
        docEdited.FirstSection.Body.Paragraphs[1].Remove();
        docEdited.StopTrackRevisions();

        // Set up comparison options. No flags are set to ignore deletions,
        // so deleted content will be part of the comparison result.
        CompareOptions compareOptions = new CompareOptions
        {
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison; revisions (including deletions) are added to docOriginal.
        docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);

        // Retain deleted text in the output by showing the original revision.
        docOriginal.LayoutOptions.RevisionOptions.ShowOriginalRevision = true;

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
