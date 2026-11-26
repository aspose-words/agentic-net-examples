using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TrackChangesSelectiveAccept
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Section 1 content.
            builder.Writeln("Section 1 - Original text.");

            // Insert a section break to start Section 2.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 - Original text.");

            // Enable track changes.
            doc.StartTrackRevisions("Demo Author", DateTime.Now);

            // Make a change in Section 1 (insertion).
            builder.MoveToDocumentStart();
            builder.Writeln("Inserted line in Section 1.");

            // Make a change in Section 2 (insertion).
            builder.MoveToDocumentEnd();
            builder.Writeln("Inserted line in Section 2.");

            // Stop tracking.
            doc.StopTrackRevisions();

            // Save the document with revisions for reference.
            doc.Save("RevisionsDocument.docx");

            // Collect revisions that belong to Section 1.
            List<Revision> revisionsToAccept = new List<Revision>();
            foreach (Revision rev in doc.Revisions)
            {
                // Find the section that contains the revision.
                Section revSection = rev.ParentNode?.GetAncestor(NodeType.Section) as Section;
                if (revSection == null)
                    continue;

                // Determine the index of the section in the document.
                int sectionIndex = doc.Sections.IndexOf(revSection);

                // Accept revisions only in the first section (index 0).
                if (sectionIndex == 0)
                    revisionsToAccept.Add(rev);
            }

            // Accept the collected revisions.
            foreach (Revision rev in revisionsToAccept)
                rev.Accept();

            // Save the resulting document.
            doc.Save("SelectiveAccepted.docx");
        }
    }
}
