using System;
using Aspose.Words;

namespace RevisionInfoExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First revision – author "Alice".
            doc.StartTrackRevisions("Alice", DateTime.Now);
            builder.Writeln("First revision text.");
            doc.StopTrackRevisions();

            // Second revision – author "Bob".
            doc.StartTrackRevisions("Bob", DateTime.Now.AddMinutes(5));
            builder.Writeln("Second revision text.");

            // Create a deletion revision by removing the first run (the text added by Alice).
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            doc.StopTrackRevisions();

            // Save the document (optional, demonstrates the save lifecycle).
            doc.Save("RevisionsDemo.docx");

            // Iterate through all revisions and log each revision's author and timestamp.
            foreach (Revision rev in doc.Revisions)
            {
                Console.WriteLine($"Author: {rev.Author}, DateTime: {rev.DateTime}");
            }
        }
    }
}
