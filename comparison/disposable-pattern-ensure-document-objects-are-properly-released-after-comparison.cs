using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            const string originalPath = @"C:\Docs\Original.docx";
            const string editedPath   = @"C:\Docs\Edited.docx";

            // Use using statements to guarantee that the file streams are closed and disposed
            // as soon as the comparison operation is finished.
            using (FileStream originalStream = File.OpenRead(originalPath))
            using (FileStream editedStream   = File.OpenRead(editedPath))
            {
                // Load the documents from the streams. The Document constructors
                // accept a Stream and therefore do not keep the stream open after loading.
                Document originalDoc = new Document(originalStream);
                Document editedDoc   = new Document(editedStream);

                // Ensure both documents have no revisions before performing a comparison.
                if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
                {
                    // Configure comparison options as needed.
                    CompareOptions compareOptions = new CompareOptions
                    {
                        // Example: ignore case changes and formatting differences.
                        IgnoreCaseChanges = true,
                        IgnoreFormatting  = true,
                        // Use the original document as the base for comparison.
                        Target = ComparisonTargetType.Current
                    };

                    // Perform the comparison. The revisions are added to originalDoc.
                    originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);
                }

                // Save the comparison result. The Save method determines the format from the file extension.
                originalDoc.Save(@"C:\Docs\ComparisonResult.docx");
            } // Both FileStream objects are disposed here; the Document instances are now eligible for GC.
        }
    }
}
