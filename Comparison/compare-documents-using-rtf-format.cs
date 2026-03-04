using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

namespace DocumentComparisonRtf
{
    class Program
    {
        static void Main()
        {
            // Paths to the source RTF files.
            string originalPath = @"C:\Docs\Original.rtf";
            string editedPath   = @"C:\Docs\Edited.rtf";

            // Load the original document with default RTF load options.
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document originalDoc = new Document(originalPath, loadOptions);

            // Load the edited document with default RTF load options.
            Document editedDoc = new Document(editedPath, loadOptions);

            // Ensure both documents have no revisions before comparison (required by Aspose.Words).
            if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
            {
                // Compare the documents. Revisions will be added to the original document.
                originalDoc.Compare(editedDoc, "JD", DateTime.Now);
            }

            // Save the comparison result as an RTF file with compact size option enabled.
            RtfSaveOptions saveOptions = new RtfSaveOptions
            {
                ExportCompactSize = true,               // Reduce file size (no RTL text in this case).
                ExportImagesForOldReaders = false       // Smaller size, older readers may not display images.
            };

            string resultPath = @"C:\Docs\ComparisonResult.rtf";
            originalDoc.Save(resultPath, saveOptions);
        }
    }
}
