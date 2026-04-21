using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text without tracking – this should not create any revisions.
        builder.Write("Initial content. ");

        // Check and display whether the document has revisions at this point.
        Console.WriteLine($"Has revisions (before tracking): {doc.HasRevisions}");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Write additional text – this will be recorded as a revision.
        builder.Write("Added revision content. ");

        // Stop tracking further changes (optional for this example).
        doc.StopTrackRevisions();

        // Check and display whether the document now contains revisions.
        Console.WriteLine($"Has revisions (after tracking): {doc.HasRevisions}");

        // Save the document to the local file system.
        doc.Save("SampleRevisions.docx");
    }
}
