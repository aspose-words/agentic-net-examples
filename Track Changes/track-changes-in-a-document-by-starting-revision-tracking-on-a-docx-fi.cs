using System;
using Aspose.Words;

namespace RevisionTrackingExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX file.
            Document doc = new Document("Input.docx");

            // Start tracking revisions. All subsequent edits will be recorded as revisions.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            // Perform some edits that will be captured as revisions.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This paragraph is added while revision tracking is enabled.");

            // Save the document with the tracked changes.
            doc.Save("Output.docx");
        }
    }
}
