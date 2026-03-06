using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class CompareRtfDocuments
{
    static void Main()
    {
        // Path to the folder that contains the RTF files.
        string dataDir = @"C:\Docs\";

        // Load the original RTF document.
        Document docOriginal = new Document(
            dataDir + "Original.rtf",
            new RtfLoadOptions());

        // Load the edited RTF document.
        Document docEdited = new Document(
            dataDir + "Edited.rtf",
            new RtfLoadOptions());

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Save the comparison result (original document with revisions) as RTF.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Reduce file size; safe here because the documents do not contain RTL text.
            ExportCompactSize = true
        };

        docOriginal.Save(dataDir + "ComparedResult.rtf", saveOptions);
    }
}
