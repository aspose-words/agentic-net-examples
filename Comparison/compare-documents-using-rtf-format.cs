using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class RtfDocumentComparison
{
    static void Main()
    {
        // Paths to the source RTF files and the folder where the result will be saved.
        string dataDir = @"C:\Docs\";
        string artifactsDir = @"C:\Output\";

        // Load the original and edited RTF documents with default load options.
        // RtfLoadOptions can be customized if needed (e.g., RecognizeUtf8Text).
        RtfLoadOptions loadOptions = new RtfLoadOptions();

        Document docOriginal = new Document(dataDir + "Original.rtf", loadOptions);
        Document docEdited   = new Document(dataDir + "Edited.rtf",   loadOptions);

        // The Compare method requires that both documents have no existing revisions.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. All differences will be stored as revisions in docOriginal.
            docOriginal.Compare(docEdited, "JD", DateTime.Now);
        }

        // Configure RTF save options.
        // ExportCompactSize reduces file size (acceptable when no RTL text is present).
        // ExportImagesForOldReaders = false keeps the file smaller by omitting extra keywords.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            ExportCompactSize = true,
            ExportImagesForOldReaders = false
        };

        // Save the comparison result as an RTF file.
        docOriginal.Save(artifactsDir + "ComparisonResult.rtf", saveOptions);
    }
}
