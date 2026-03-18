using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class RevisionDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few paragraphs that we will later change the style of.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        // -----------------------------------------------------------------
        // 1. Enable tracking of changes.
        // -----------------------------------------------------------------
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // -----------------------------------------------------------------
        // 2. Apply a style change to all existing paragraphs while tracking is on.
        //    (Note: formatting changes are not recorded as revisions, but the
        //    operation is performed as requested.)
        // -----------------------------------------------------------------
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            // Change the paragraph style to "Heading 1".
            para.ParagraphFormat.StyleName = "Heading 1";
        }

        // -----------------------------------------------------------------
        // 3. Insert a new paragraph – this will generate an insertion revision.
        // -----------------------------------------------------------------
        builder.Writeln("This paragraph is inserted while tracking is enabled.");

        // -----------------------------------------------------------------
        // 4. Stop tracking so subsequent edits are not recorded as revisions.
        // -----------------------------------------------------------------
        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 5. Verify that there is exactly one revision group (the insertion above).
        // -----------------------------------------------------------------
        if (doc.Revisions.Groups.Count == 1)
        {
            RevisionGroup group = doc.Revisions.Groups[0];
            Console.WriteLine($"Revision group verified:");
            Console.WriteLine($"  Author: {group.Author}");
            Console.WriteLine($"  Type:   {group.RevisionType}");
            Console.WriteLine($"  Text:   {group.Text.Trim()}");
        }
        else
        {
            Console.WriteLine($"Unexpected number of revision groups: {doc.Revisions.Groups.Count}");
        }

        // Save the document (using the standard save method).
        doc.Save("RevisionDemo.docx");
    }
}
