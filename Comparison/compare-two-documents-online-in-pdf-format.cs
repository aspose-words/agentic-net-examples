using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

class CompareOnlineDocumentsToPdf
{
    static void Main()
    {
        // URLs of the two source documents (DOCX format recommended for comparison)
        const string urlOriginal = "https://filesamples.com/samples/document/docx/sample1.docx";
        const string urlEdited   = "https://filesamples.com/samples/document/docx/sample2.docx";

        // Download the documents into memory streams
        using (HttpClient httpClient = new HttpClient())
        {
            // Original document
            byte[] originalBytes = httpClient.GetByteArrayAsync(urlOriginal).Result;
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                Document docOriginal = new Document(originalStream);

                // Edited document
                byte[] editedBytes = httpClient.GetByteArrayAsync(urlEdited).Result;
                using (MemoryStream editedStream = new MemoryStream(editedBytes))
                {
                    Document docEdited = new Document(editedStream);

                    // Ensure both documents have no revisions before comparison
                    if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
                    {
                        // Compare the edited document against the original.
                        // Revisions will be added to docOriginal indicating the differences.
                        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
                    }

                    // Save the comparison result (with revisions) as a PDF.
                    // The PDF will display tracked changes similar to Word's "Show changes".
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        // Optional: keep the revisions visible in the PDF.
                        // This setting ensures that the PDF contains the revision marks.
                        UpdateFields = true
                    };

                    docOriginal.Save("ComparisonResult.pdf", pdfOptions);
                }
            }
        }
    }
}
