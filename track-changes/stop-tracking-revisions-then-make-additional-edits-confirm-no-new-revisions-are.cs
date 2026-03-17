using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin tracking revisions.
        doc.StartTrackRevisions("John Doe");

        // First edit – will be recorded as a revision.
        builder.Writeln("First revision text.");

        // Capture revision count after the first edit.
        int countAfterFirstEdit = doc.Revisions.Count;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Second edit – should NOT create a new revision.
        builder.Writeln("Text after tracking stopped.");

        // Capture revision count after the second edit.
        int countAfterSecondEdit = doc.Revisions.Count;

        // Verify that no new revisions were added.
        Console.WriteLine($"Revisions after first edit: {countAfterFirstEdit}");
        Console.WriteLine($"Revisions after second edit: {countAfterSecondEdit}");
        Console.WriteLine($"No new revisions recorded: {countAfterSecondEdit == countAfterFirstEdit}");

        // Save the document.
        doc.Save("RevisionsDemo.docx");
    }
}
