using System;
using Aspose.Words;

class RevisionLogger
{
    static void Main()
    {
        // Load an existing Word document.
        // Replace the path with the actual location of your document.
        Document doc = new Document("Input.docx");

        // Iterate through all revisions in the document.
        foreach (Revision revision in doc.Revisions)
        {
            // Log the author and the date/time of each revision.
            Console.WriteLine($"Author: {revision.Author}, Timestamp: {revision.DateTime}");
        }

        // Optionally, save the document (unchanged) to a new file.
        // This demonstrates the required save lifecycle step.
        doc.Save("Output.docx");
    }
}
