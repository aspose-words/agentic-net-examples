using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable track changes with a specific author and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a new paragraph (creates an insertion revision).
        builder.Writeln("Inserted paragraph.");

        // Delete the original paragraph (creates a deletion revision).
        Node originalParagraph = doc.FirstSection.Body.FirstParagraph;
        originalParagraph.Remove();

        // Change formatting of the inserted paragraph (creates a format change revision).
        Node insertedParagraph = doc.FirstSection.Body.FirstParagraph;
        if (insertedParagraph is Paragraph para)
        {
            para.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates that revisions are persisted).
        doc.Save("Revisions.docx");

        // Iterate through all revisions and log each revision's author and timestamp.
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Author: {rev.Author}, Timestamp: {rev.DateTime}");
        }
    }
}
