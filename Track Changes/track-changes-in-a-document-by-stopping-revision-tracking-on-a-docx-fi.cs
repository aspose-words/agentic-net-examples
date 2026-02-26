using System;
using Aspose.Words;

class StopRevisionTrackingExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Stop automatic tracking of revisions. Any further edits will not be recorded as revisions.
        doc.StopTrackRevisions();

        // Save the document after stopping revision tracking.
        doc.Save("OutputDocument.docx");
    }
}
