using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new blank document and a builder.
        // -------------------------------------------------
        Document doc = new Document();                 // create blank document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text – this will NOT be tracked.
        builder.Writeln("Paragraph before tracking.");

        // -------------------------------------------------
        // 2. Start tracking revisions.
        // -------------------------------------------------
        doc.StartTrackRevisions("Alice");               // author name for revisions

        // Insert a new paragraph – recorded as an insertion revision.
        builder.Writeln("This line is inserted while tracking.");

        // Modify the first paragraph: delete its original run and insert a new one.
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        firstPara.Runs[0].Remove();                    // deletion revision
        firstPara.Runs.Add(new Run(doc, "Modified first paragraph.")); // insertion revision

        // -------------------------------------------------
        // 3. Stop tracking revisions.
        // -------------------------------------------------
        doc.StopTrackRevisions();

        // Add more content – this will NOT be tracked.
        builder.Writeln("Paragraph after tracking.");

        // -------------------------------------------------
        // 4. Save the document containing revisions.
        // -------------------------------------------------
        doc.Save("TrackedChanges.docx");                // save with revisions

        // -------------------------------------------------
        // 5. Accept all revisions programmatically and save a clean copy.
        // -------------------------------------------------
        doc.AcceptAllRevisions();                      // remove all revisions
        doc.Save("TrackedChanges_Accepted.docx");      // clean version

        // -------------------------------------------------
        // 6. Demonstrate document comparison which also creates revisions.
        // -------------------------------------------------
        Document original = new Document("TrackedChanges.docx");
        Document edited   = new Document("TrackedChanges_Accepted.docx");

        // Ensure both documents are revision‑free before comparison.
        if (original.HasRevisions) original.Revisions.AcceptAll();
        if (edited.HasRevisions)   edited.Revisions.AcceptAll();

        // Compare the two documents; revisions will be added to 'original'.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Save the comparison result – revisions are visible in Word.
        original.Save("ComparisonResult.docx");
    }
}
