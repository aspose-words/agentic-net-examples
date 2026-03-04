using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class RtfDocumentComparison
{
    static void Main()
    {
        // Path to the folder that contains the source RTF files.
        string docsPath = @"C:\Docs\";

        // Load the original and edited documents using RtfLoadOptions.
        // The load options can be customized if needed (e.g., encoding, recovery mode).
        Document originalDoc = new Document(docsPath + "Original.rtf", new RtfLoadOptions());
        Document editedDoc   = new Document(docsPath + "Edited.rtf",   new RtfLoadOptions());

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            // Compare the documents. The original document will receive Revision objects
            // that describe the differences found in the edited document.
            originalDoc.Compare(editedDoc, "JD", DateTime.Now);
        }

        // Prepare RTF save options. For example, enable compact size to reduce file size
        // (acceptable when the document does not contain right‑to‑left text).
        RtfSaveOptions rtfSaveOptions = new RtfSaveOptions
        {
            ExportCompactSize = true
        };

        // Save the comparison result back to RTF format.
        originalDoc.Save(docsPath + "ComparedResult.rtf", rtfSaveOptions);
    }
}
