using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Build initial content ----------
        // Section 1
        builder.Writeln("Original paragraph in Section 1.");
        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // Section 2
        builder.Writeln("Original paragraph in Section 2.");

        // Keep references to the two sections.
        Section section1 = doc.FirstSection;
        Section section2 = doc.Sections[1];

        // ---------- Create revisions in Section 1 ----------
        doc.StartTrackRevisions("AuthorA", DateTime.Now);

        // Insert a new paragraph at the beginning of Section 1.
        builder.MoveTo(section1.Body.FirstParagraph);
        builder.Writeln("Inserted revision in Section 1.");

        // Delete the original paragraph in Section 1 to generate a deletion revision.
        Node paragraphToDelete = section1.Body.Paragraphs[1]; // the original paragraph
        paragraphToDelete.Remove();

        doc.StopTrackRevisions();

        // ---------- Create revisions in Section 2 ----------
        doc.StartTrackRevisions("AuthorB", DateTime.Now);

        // Insert a new paragraph at the end of Section 2.
        builder.MoveTo(section2.Body.LastParagraph);
        builder.Writeln("Inserted revision in Section 2.");

        // Delete the original paragraph in Section 2 to generate a deletion revision.
        Node paraToDelete = section2.Body.Paragraphs[0]; // the original paragraph
        paraToDelete.Remove();

        doc.StopTrackRevisions();

        // ---------- Accept only revisions that belong to Section 1 ----------
        List<Revision> revisionsToAccept = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            // Some revisions (e.g., style changes) may have a null ParentNode.
            if (rev.ParentNode == null) continue;

            // Find the section that contains the revision's parent node.
            Node ancestorSection = rev.ParentNode.GetAncestor(NodeType.Section);
            if (ancestorSection == section1)
                revisionsToAccept.Add(rev);
        }

        // Accept the collected revisions.
        foreach (Revision rev in revisionsToAccept)
            rev.Accept();

        // Save the resulting document.
        doc.Save("RevisionsSelectiveAccept.docx");
    }
}
