using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class RevisionComparisonToXps
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Perform a comparison. The original document will receive Revision objects.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Configure how revisions are displayed during layout.
        // Hide revision bars, keep revision marks (e.g., underline/strike‑through).
        docOriginal.LayoutOptions.RevisionOptions.ShowRevisionBars = false;
        docOriginal.LayoutOptions.RevisionOptions.ShowRevisionMarks = true;
        // Do not show revisions in balloons.
        docOriginal.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.None;

        // Set the view to the revised (final) version so that the saved XPS reflects the changes.
        docOriginal.RevisionsView = RevisionsView.Final;

        // Create XPS save options. Adjust any additional options if needed.
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // Enable high‑quality rendering for better visual fidelity.
            UseHighQualityRendering = true,
            // Embed the generator name (default true) – can be left as is.
            ExportGeneratorName = true
        };

        // Save the document with revisions applied to XPS format.
        docOriginal.Save("ComparisonResult.xps", xpsOptions);
    }
}
