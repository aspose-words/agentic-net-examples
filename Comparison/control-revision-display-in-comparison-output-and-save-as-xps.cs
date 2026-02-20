using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class RevisionComparisonToXps
{
    static void Main()
    {
        // Load the original and edited documents.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Set up comparison options – use the edited document as the target.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison; revisions are added to the original document.
        original.Compare(edited, "Reviewer", DateTime.Now, compareOptions);

        // Configure how revisions are displayed in the layout.
        RevisionOptions revOptions = original.LayoutOptions.RevisionOptions;
        revOptions.ShowInBalloons = ShowInBalloons.Format;   // Show format revisions in balloons.
        revOptions.ShowOriginalRevision = true;              // Show original text alongside revisions.
        revOptions.ShowRevisionMarks = true;                // Render revision marks (e.g., underline, strike‑through).

        // Save the resulting document as XPS using XpsSaveOptions.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        original.Save("ComparisonResult.xps", xpsOptions);
    }
}
