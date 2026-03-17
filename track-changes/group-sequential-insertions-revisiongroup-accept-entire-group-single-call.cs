using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin tracking changes – all subsequent edits will be recorded as revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert several paragraphs. These will be recorded as sequential insertion revisions.
        builder.Writeln("First line.");
        builder.Writeln("Second line.");
        builder.Writeln("Third line.");

        // Stop tracking – further edits will not be recorded.
        doc.StopTrackRevisions();

        // Iterate over all revision groups in the document.
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            // We are interested only in groups that consist of insertion revisions.
            if (group.RevisionType == RevisionType.Insertion)
            {
                // Accept every revision that belongs to this insertion group.
                foreach (Revision rev in doc.Revisions)
                {
                    if (rev.Group == group)
                        rev.Accept();
                }
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
