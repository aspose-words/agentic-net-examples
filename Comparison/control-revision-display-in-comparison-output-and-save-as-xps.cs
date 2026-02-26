using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;

// Load the original and edited documents.
string docsPath = @"C:\Docs\";
Document docOriginal = new Document(docsPath + "Original.docx");
Document docEdited   = new Document(docsPath + "Edited.docx");

// Ensure both documents have no revisions before comparison.
if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
{
    // Compare the documents – revisions will be added to docOriginal.
    docOriginal.Compare(docEdited, "Author", DateTime.Now);
}

// Show the revised version (with tracked changes) in the layout.
docOriginal.RevisionsView = RevisionsView.Final;

// Optional: customize how revisions are rendered.
docOriginal.LayoutOptions.RevisionOptions.ShowRevisionBars = true;
docOriginal.LayoutOptions.RevisionOptions.RevisionBarsPosition = HorizontalAlignment.Outside;
docOriginal.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;

// Save the comparison result as XPS.
XpsSaveOptions xpsOptions = new XpsSaveOptions();
xpsOptions.OptimizeOutput = false; // keep original layout fidelity
docOriginal.Save(docsPath + "ComparisonResult.xps", xpsOptions);
