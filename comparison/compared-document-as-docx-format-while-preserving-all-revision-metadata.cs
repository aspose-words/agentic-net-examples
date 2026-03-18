using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsRevisionSample
{
    class Program
    {
        static void Main()
        {
            // Paths to the original and edited documents.
            // In a real scenario replace these with actual file locations.
            string originalPath = @"C:\Docs\Original.docx";
            string editedPath   = @"C:\Docs\Edited.docx";

            // Load the two documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Ensure both documents have no revisions before comparison.
            // This is required by the Compare method.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
            {
                // Compare the edited document against the original.
                // The revisions will be added to docOriginal.
                docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
            }

            // Save the result as DOCX while preserving all revision metadata.
            // The Save method automatically determines the format from the extension.
            string resultPath = @"C:\Docs\ComparedResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
