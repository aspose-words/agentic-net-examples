using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

namespace AsposeWordsCompareTables
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            // docOriginal is a legacy .doc file, docEdited is a .docx file.
            string docPath = @"C:\Docs\Original.doc";
            string editedPath = @"C:\Docs\Edited.docx";

            // Load the documents using the Document constructors (lifecycle rule).
            Document docOriginal = new Document(docPath);
            Document docEdited = new Document(editedPath);

            // Configure comparison options.
            // Ensure that table differences are NOT ignored.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreTables = false,               // Detect table structure changes.
                IgnoreFormatting = false,           // Include formatting changes.
                IgnoreComments = false,             // Include comment changes.
                IgnoreHeadersAndFooters = false,    // Include header/footer changes.
                Target = ComparisonTargetType.New   // Use the edited document as the target.
            };

            // Perform the comparison. The original document will receive Revision objects.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Count revisions that are related to tables.
            int tableRevisions = 0;
            foreach (Revision rev in docOriginal.Revisions)
            {
                // The ParentNode of a revision indicates the node that changed.
                if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Table)
                    tableRevisions++;
            }

            // Output the result – a non‑zero count means table differences were detected.
            Console.WriteLine($"Total revisions: {docOriginal.Revisions.Count}");
            Console.WriteLine($"Table‑related revisions: {tableRevisions}");

            // Save the comparison result (lifecycle rule).
            string resultPath = @"C:\Docs\ComparisonResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
