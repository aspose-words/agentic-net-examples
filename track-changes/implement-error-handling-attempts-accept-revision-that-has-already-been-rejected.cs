using System;
using Aspose.Words;

class RevisionErrorHandlingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Add some text – this will be recorded as an insertion revision.
        builder.Write("This is a tracked revision.");

        // Retrieve the first revision from the collection.
        Revision revision = doc.Revisions[0];

        // Reject the revision – it is now removed from the document.
        revision.Reject();

        // Attempt to accept the same revision after it has been rejected.
        // The Revision object is no longer valid, so Accept() may throw an exception.
        try
        {
            revision.Accept();
            Console.WriteLine("Revision accepted successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Handle the error – typically an InvalidOperationException is thrown.
            Console.WriteLine($"Error: Cannot accept a revision that has already been rejected. Details: {ex.Message}");
        }

        // Save the document (it will contain no revisions after the reject).
        doc.Save("RevisionErrorHandlingOutput.docx");
    }
}
