using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Build the initial document ----------
        // Section 1 – original content.
        builder.Writeln("Section 1 original text.");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – original content.
        builder.Writeln("Section 2 original text.");

        // ---------- Create revisions in both sections ----------
        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Add a revision to Section 1.
        builder.MoveToDocumentStart(); // Position at the beginning of the document.
        builder.Writeln("Section 1 added revision.");

        // Add a revision to Section 2.
        builder.MoveToDocumentEnd(); // Position at the end of the document.
        builder.Writeln("Section 2 added revision.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document has revisions in both sections.
        // ---------- Accept revisions only in Section 1 ----------
        // Accept all revisions that belong to the first section.
        doc.FirstSection.Range.Revisions.AcceptAll();

        // The revisions in Section 2 remain untouched.

        // Save the resulting document.
        doc.Save("Output.docx");

        // Optional: output revision counts to the console for verification.
        Console.WriteLine($"Total revisions after selective accept: {doc.Revisions.Count}");
        Console.WriteLine($"Revisions in Section 1: {doc.FirstSection.Range.Revisions.Count}");
        Console.WriteLine($"Revisions in Section 2: {doc.LastSection.Range.Revisions.Count}");
    }
}
