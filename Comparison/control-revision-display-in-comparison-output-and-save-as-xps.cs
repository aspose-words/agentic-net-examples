using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the original and edited documents.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(edited, "Author", DateTime.Now);

        // Customize how revisions are displayed in the layout.
        // For example, hide revision bars and set the view to the final (revised) version.
        original.LayoutOptions.RevisionOptions.ShowRevisionBars = false;
        original.RevisionsView = RevisionsView.Final;

        // Create XPS save options. Enable output optimization if desired.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.OptimizeOutput = true;

        // Save the comparison result as an XPS file.
        original.Save("ComparisonResult.xps", xpsOptions);
    }
}
