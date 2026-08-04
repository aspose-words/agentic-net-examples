using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Initial content. ");

        // First author makes a revision.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Writeln("John's revision text. ");
        doc.StopTrackRevisions();

        // Second author makes a revision.
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Writeln("Jane's revision text. ");
        doc.StopTrackRevisions();

        // At this point the document has two insertion revisions.
        Console.WriteLine($"Revisions before rejection: {doc.Revisions.Count}");

        // Reject all revisions authored by "Jane Smith".
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.Author == "Jane Smith")
                rev.Reject();
        }

        // Verify that no revision from Jane Smith remains.
        bool janeRevisionExists = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author == "Jane Smith")
            {
                janeRevisionExists = true;
                break;
            }
        }

        Console.WriteLine($"Revisions after rejection: {doc.Revisions.Count}");
        Console.WriteLine($"Jane's revision still present: {janeRevisionExists}");

        // Confirm that the text added by Jane is no longer in the document.
        string docText = doc.GetText();
        bool janeTextPresent = docText.Contains("Jane's revision text");
        Console.WriteLine($"Jane's revision text present in document: {janeTextPresent}");

        // Save the resulting document.
        doc.Save("RevisionsResult.docx");
    }
}
