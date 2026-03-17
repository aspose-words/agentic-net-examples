using System;
using Aspose.Words;

class TrackRevisionsToggle
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content.
        builder.Writeln("Initial content. ");

        Console.WriteLine("Type 'on' to start tracking revisions, 'off' to stop tracking, or 'exit' to finish.");

        while (true)
        {
            Console.Write("> ");
            string input = Console.ReadLine()?.Trim().ToLowerInvariant();

            if (input == "exit")
                break;

            if (input == "on")
            {
                // Start tracking programmatic changes.
                doc.StartTrackRevisions("ConsoleUser", DateTime.Now);
                // Also set the Word UI flag (optional, affects Word UI only).
                doc.TrackRevisions = true;

                // Example change that will be recorded as a revision.
                builder.Writeln("Added while tracking is ON. ");
                Console.WriteLine("Tracking started. Added a line.");
            }
            else if (input == "off")
            {
                // Stop tracking programmatic changes.
                doc.StopTrackRevisions();
                // Reset the Word UI flag.
                doc.TrackRevisions = false;

                // Example change that will NOT be recorded as a revision.
                builder.Writeln("Added while tracking is OFF. ");
                Console.WriteLine("Tracking stopped. Added a line.");
            }
            else
            {
                Console.WriteLine("Unrecognized command. Use 'on', 'off', or 'exit'.");
            }
        }

        // Save the document to disk.
        doc.Save("TrackedDocument.docx");
        Console.WriteLine("Document saved as 'TrackedDocument.docx'.");
    }
}
