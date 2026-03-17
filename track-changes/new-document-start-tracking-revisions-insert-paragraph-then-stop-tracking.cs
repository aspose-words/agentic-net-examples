using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin tracking revisions. All subsequent changes will be recorded as revisions.
        doc.StartTrackRevisions("John Doe");

        // Insert a new paragraph with some text. This paragraph will be marked as an insertion revision.
        builder.Writeln("This is a tracked paragraph.");

        // Stop tracking revisions. Further changes will not be recorded.
        doc.StopTrackRevisions();

        // Save the document to a file.
        doc.Save("TrackedRevisions.docx");
    }
}
