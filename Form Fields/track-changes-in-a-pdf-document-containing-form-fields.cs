using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class TrackChangesInPdfWithFormFields
{
    static void Main()
    {
        // Load a Word document that already contains form fields.
        Document doc = new Document("FormFields.docx");

        // Begin tracking revisions programmatically.
        doc.StartTrackRevisions("Reviewer");

        // Example modification: change the result of the first form field.
        // This change will be recorded as an insertion revision.
        FormField firstField = doc.Range.FormFields[0];
        firstField.Result = "Updated value";

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Prepare PDF save options to preserve the form fields as interactive PDF fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,
            // Optional: keep the original author name on the PDF revision annotations.
            // The PDF format does not store revision authors directly, but the
            // appearance of tracked changes will reflect the revisions made above.
        };

        // Save the document as PDF. The tracked changes will be visible in the PDF.
        doc.Save("TrackedFormFields.pdf", pdfOptions);
    }
}
