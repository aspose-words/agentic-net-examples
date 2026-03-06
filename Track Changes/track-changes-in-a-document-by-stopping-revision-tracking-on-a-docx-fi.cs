using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Stop automatic tracking of revisions.
        doc.StopTrackRevisions();

        // Save the document with revisions tracking disabled.
        doc.Save("Output.docx");
    }
}
