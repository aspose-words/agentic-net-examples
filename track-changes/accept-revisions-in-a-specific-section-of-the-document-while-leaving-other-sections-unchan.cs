using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build two sections with some initial content.
        builder.Writeln("Section 1");
        builder.Writeln("Original text in section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.Writeln("Original text in section 2.");

        // Enable track changes.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert a paragraph in Section 1 (revision will be tracked).
        builder.MoveTo(doc.Sections[0].Body.LastParagraph);
        builder.Writeln("Inserted revision in Section 1.");

        // Insert a paragraph in Section 2 (revision will be tracked).
        builder.MoveTo(doc.Sections[1].Body.LastParagraph);
        builder.Writeln("Inserted revision in Section 2.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Accept revisions only in Section 1, leave others unchanged.
        Section targetSection = doc.Sections[0];

        // Copy revisions to a list because we will modify the collection while iterating.
        List<Revision> revisions = doc.Revisions.Cast<Revision>().ToList();

        foreach (Revision rev in revisions)
        {
            // Find the section that contains this revision.
            Section revSection = rev.ParentNode.GetAncestor(NodeType.Section) as Section;
            if (revSection == targetSection)
            {
                rev.Accept();
            }
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
