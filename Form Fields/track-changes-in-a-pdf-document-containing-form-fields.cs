using System;
using Aspose.Words;
using Aspose.Words.Saving;

class TrackChangesPdfWithFormFields
{
    static void Main()
    {
        // Load an existing Word document that contains form fields.
        // (Assumes the file "FormFields.docx" exists in the working directory.)
        Document doc = new Document("FormFields.docx");

        // Start tracking revisions programmatically.
        // All subsequent changes will be recorded as revision changes.
        doc.StartTrackRevisions("Automated", DateTime.Now);

        // Make a sample change – add a paragraph after the first one.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.Writeln("This paragraph was added while tracking revisions.");

        // Stop tracking revisions so further edits are not recorded.
        doc.StopTrackRevisions();

        // Configure PDF save options to preserve the Word form fields as interactive PDF form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };

        // Save the document as PDF with the specified options.
        doc.Save("TrackedChanges.pdf", pdfOptions);
    }
}
