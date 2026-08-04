using System;
using Aspose.Words;

public class AcceptRevisionsInSection
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content to the first section.
        builder.Writeln("Section 1 original text.");

        // Insert a section break to start the second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add content to the second section.
        builder.Writeln("Section 2 original text.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Add a revision to the first section.
        builder.MoveToSection(0);
        builder.Writeln("Revision added to Section 1.");

        // Add a revision to the second section.
        builder.MoveToSection(1);
        builder.Writeln("Revision added to Section 2.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all revisions only in the first section.
        doc.Sections[0].Range.Revisions.AcceptAll();

        // Save the resulting document.
        doc.Save("AcceptRevisionsInSection.docx");
    }
}
