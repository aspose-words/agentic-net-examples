using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Stop tracking revisions (track changes) in the document.
        doc.StopTrackRevisions();

        // Save the document with revisions tracking disabled.
        doc.Save("Output.docx");
    }
}
