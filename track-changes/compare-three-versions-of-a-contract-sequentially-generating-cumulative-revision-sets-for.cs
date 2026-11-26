using System;
using Aspose.Words;

namespace TrackChangesComparison
{
    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // Step 1: Create the initial contract (Version 1).
            // -------------------------------------------------
            Document docV1 = new Document();
            DocumentBuilder builderV1 = new DocumentBuilder(docV1);
            builderV1.Writeln("Contract");
            builderV1.Writeln("Version 1");
            builderV1.Writeln("This contract is between Party A and Party B.");
            docV1.Save("Contract_v1.docx");

            // -------------------------------------------------
            // Step 2: Load Version 1, enable tracking, modify to create Version 2.
            // -------------------------------------------------
            Document docV2 = new Document("Contract_v1.docx");
            docV2.StartTrackRevisions("Author1", DateTime.Now);
            DocumentBuilder builderV2 = new DocumentBuilder(docV2);

            // Insert a new version line at the beginning.
            builderV2.MoveToDocumentStart();
            builderV2.Writeln("Version 2");

            // Add a new clause.
            builderV2.Writeln("Added clause: Confidentiality.");

            docV2.StopTrackRevisions();
            docV2.Save("Contract_v2.docx");

            // Inspect revisions that represent the changes from V1 to V2.
            RevisionCollection revsV2 = docV2.Revisions;
            if (revsV2.Count == 0)
                throw new InvalidOperationException("No revisions were generated between Version 1 and Version 2.");

            Console.WriteLine($"Revisions between Version 1 and Version 2: {revsV2.Count}");
            foreach (Revision rev in revsV2)
            {
                Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            }

            // -------------------------------------------------
            // Step 3: Load Version 2, enable tracking, modify to create Version 3.
            // -------------------------------------------------
            Document docV3 = new Document("Contract_v2.docx");
            docV3.StartTrackRevisions("Author2", DateTime.Now);
            DocumentBuilder builderV3 = new DocumentBuilder(docV3);

            // Insert a new version line at the beginning.
            builderV3.MoveToDocumentStart();
            builderV3.Writeln("Version 3");

            // Update the confidentiality clause.
            builderV3.Writeln("Updated clause: Confidentiality extended.");

            docV3.StopTrackRevisions();
            docV3.Save("Contract_v3.docx");

            // Inspect cumulative revisions present in Version 3 (includes V1→V2 and V2→V3).
            RevisionCollection revsV3 = docV3.Revisions;
            if (revsV3.Count == 0)
                throw new InvalidOperationException("No revisions were generated in Version 3.");

            Console.WriteLine($"Cumulative revisions in Version 3 (V1→V2 and V2→V3): {revsV3.Count}");
            foreach (Revision rev in revsV3)
            {
                Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            }

            // Optional: Accept all revisions to finalize the document.
            docV3.AcceptAllRevisions();
            docV3.Save("Contract_v3_final.docx");
        }
    }
}
