using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path to the destination PDF file.
            const string outputPath = @"C:\Docs\ResultDocument.pdf";

            // Load the existing DOCX document using the Document(string) constructor.
            Document doc = new Document(inputPath);

            // -----------------------------------------------------------------
            // Customize text properties: change font name and size for every Run.
            // -----------------------------------------------------------------
            // Get all Run nodes in the document (including those inside tables, headers, etc.).
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                // Set the desired font name and size.
                run.Font.Name = "Arial";
                run.Font.Size = 12; // Font size in points.
            }

            // ---------------------------------------------------------------
            // Save the modified document as PDF.
            // ---------------------------------------------------------------
            // Create PdfSaveOptions to demonstrate usage of a save options object.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example option: write additional text positioning operators.
                AdditionalTextPositioning = false,
                // Ensure fields are updated before saving.
                UpdateFields = true
            };

            // Save the document. The overload with SaveOptions respects the PDF format.
            doc.Save(outputPath, pdfOptions);
        }
    }
}
