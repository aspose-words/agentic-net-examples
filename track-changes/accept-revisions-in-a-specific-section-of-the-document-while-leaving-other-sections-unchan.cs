using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Section 1 -----
        builder.Writeln("Section 1 - Original text.");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- Section 2 -----
        builder.Writeln("Section 2 - Original text.");

        // ---------- Create revisions in Section 1 ----------
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Move the builder to Section 1 and add a new paragraph (revision).
        builder.MoveToSection(0);
        builder.Writeln("Added in Section 1.");

        // Delete the first run of the first paragraph in Section 1 (revision).
        Paragraph firstParaSec1 = doc.Sections[0].Body.FirstParagraph;
        if (firstParaSec1.Runs.Count > 0)
            firstParaSec1.Runs[0].Remove();

        doc.StopTrackRevisions();

        // ---------- Create revisions in Section 2 ----------
        doc.StartTrackRevisions("Author2", DateTime.Now);

        // Move the builder to Section 2 and add a new paragraph (revision).
        builder.MoveToSection(1);
        builder.Writeln("Added in Section 2.");

        // Delete the first run of the first paragraph in Section 2 (revision).
        Paragraph firstParaSec2 = doc.Sections[1].Body.FirstParagraph;
        if (firstParaSec2.Runs.Count > 0)
            firstParaSec2.Runs[0].Remove();

        doc.StopTrackRevisions();

        // ---------- Accept only the revisions that belong to Section 1 ----------
        // Collect revisions from Section 1 into a list to avoid modifying the collection while iterating.
        var revisionsInSection1 = doc.Sections[0].Range.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisionsInSection1)
        {
            rev.Accept();
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
