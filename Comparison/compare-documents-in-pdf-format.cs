using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the PDF documents to compare.
            string originalPdfPath = @"C:\Docs\Original.pdf";
            string editedPdfPath   = @"C:\Docs\Edited.pdf";

            // Load the two PDF files as Aspose.Words Document objects.
            Document docOriginal = new Document(originalPdfPath);
            Document docEdited   = new Document(editedPdfPath);

            // Ensure both documents have no revisions before performing the comparison.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
            {
                // Compare the documents. The revisions will be added to docOriginal.
                docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
            }

            // Optional: iterate over the revisions to see what changed.
            foreach (Revision rev in docOriginal.Revisions)
            {
                Console.WriteLine($"Revision type: {rev.RevisionType}, Node type: {rev.ParentNode.NodeType}");
                Console.WriteLine($"\tChanged text: \"{rev.ParentNode.GetText().Trim()}\"");
            }

            // Save the comparison result (original document with revisions) as a PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set PDF compliance level if required.
                Compliance = PdfCompliance.Pdf17
            };

            string resultPdfPath = @"C:\Docs\ComparisonResult.pdf";
            docOriginal.Save(resultPdfPath, pdfOptions);

            Console.WriteLine($"Comparison PDF saved to: {resultPdfPath}");
        }
    }
}
