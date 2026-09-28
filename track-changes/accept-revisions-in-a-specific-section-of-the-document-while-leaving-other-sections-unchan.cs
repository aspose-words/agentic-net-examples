using System;
using System.Collections.Generic;
using Aspose.Words;

namespace TrackChangesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build Section 1.
            builder.Writeln("Section 1 - Original text.");
            // Insert a section break to start Section 2.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            // Build Section 2.
            builder.Writeln("Section 2 - Original text.");

            // Enable tracking of revisions.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // Add a revision (insertion) in Section 1.
            Paragraph paraSec1 = new Paragraph(doc);
            paraSec1.AppendChild(new Run(doc, "Inserted revision in Section 1."));
            doc.Sections[0].Body.AppendChild(paraSec1);

            // Add a revision (insertion) in Section 2.
            Paragraph paraSec2 = new Paragraph(doc);
            paraSec2.AppendChild(new Run(doc, "Inserted revision in Section 2."));
            doc.Sections[1].Body.AppendChild(paraSec2);

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Accept only revisions that belong to Section 1.
            Section targetSection = doc.Sections[0];
            // Collect revisions to accept first to avoid modifying the collection during enumeration.
            List<Revision> revisionsToAccept = new List<Revision>();
            foreach (Revision rev in doc.Revisions)
            {
                Node ancestorSection = rev.ParentNode.GetAncestor(NodeType.Section);
                if (ancestorSection == targetSection)
                {
                    revisionsToAccept.Add(rev);
                }
            }

            // Accept the collected revisions.
            foreach (Revision rev in revisionsToAccept)
            {
                rev.Accept();
            }

            // Save the resulting document.
            doc.Save("Output.docx");
        }
    }
}
