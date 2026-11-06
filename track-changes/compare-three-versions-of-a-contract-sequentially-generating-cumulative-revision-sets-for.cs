using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class TrackChangesComparison
{
    public static void Main()
    {
        // Create the original contract (Version 1).
        Document version1 = CreateContractVersion1();
        version1.Save("Version1.docx");

        // Create Version 2 by applying first set of changes with tracking.
        Document version2 = (Document)version1.Clone();
        ApplyFirstChanges(version2, "Alice");
        version2.Save("Version2.docx");
        RevisionCollection revs12 = version2.Revisions;
        PrintRevisions("Revisions from Version 1 to Version 2", revs12);

        // Create Version 3 by applying second set of changes with tracking on Version 2.
        Document version3 = (Document)version2.Clone();
        ApplySecondChanges(version3, "Bob");
        version3.Save("Version3.docx");
        RevisionCollection revs23 = version3.Revisions;
        PrintRevisions("Revisions from Version 2 to Version 3", revs23);

        // Generate cumulative revisions from Version 1 directly to Version 3.
        Document cumulative = (Document)version1.Clone();
        ApplyFirstChanges(cumulative, "Alice");
        ApplySecondChanges(cumulative, "Bob");
        cumulative.Save("Cumulative.docx");
        RevisionCollection revs13 = cumulative.Revisions;
        PrintRevisions("Cumulative revisions from Version 1 to Version 3", revs13);
    }

    // Creates the initial contract document.
    private static Document CreateContractVersion1()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contract Title");
        builder.Writeln("Clause 1: The buyer shall pay $1000.");
        builder.Writeln("Clause 2: Delivery shall be within 30 days.");
        return doc;
    }

    // Applies the first set of modifications (Version 1 -> Version 2) while tracking revisions.
    private static void ApplyFirstChanges(Document doc, string author)
    {
        doc.StartTrackRevisions(author, DateTime.Now);

        // Modify Clause 1 text.
        Paragraph clause1 = FindParagraphContaining(doc, "Clause 1:");
        if (clause1 != null)
        {
            // Remove existing runs and insert new text as a single run.
            clause1.Runs.Clear();
            clause1.AppendChild(new Run(doc, "Clause 1: The buyer shall pay $1200."));
        }

        // Add a new clause.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.Writeln("Clause 3: Warranty period is 12 months.");

        doc.StopTrackRevisions();
    }

    // Applies the second set of modifications (Version 2 -> Version 3) while tracking revisions.
    private static void ApplySecondChanges(Document doc, string author)
    {
        doc.StartTrackRevisions(author, DateTime.Now);

        // Delete Clause 2.
        Paragraph clause2 = FindParagraphContaining(doc, "Clause 2:");
        if (clause2 != null)
        {
            clause2.Remove();
        }

        // Make Clause 1 bold.
        Paragraph clause1 = FindParagraphContaining(doc, "Clause 1:");
        if (clause1 != null)
        {
            foreach (Run run in clause1.GetChildNodes(NodeType.Run, true))
            {
                run.Font.Bold = true;
            }
        }

        doc.StopTrackRevisions();
    }

    // Finds the first paragraph that contains the specified text.
    private static Paragraph FindParagraphContaining(Document doc, string text)
    {
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains(text))
                return para;
        }
        return null;
    }

    // Prints revision details to the console.
    private static void PrintRevisions(string title, RevisionCollection revisions)
    {
        Console.WriteLine(title);
        Console.WriteLine(new string('-', title.Length));
        if (revisions.Count == 0)
        {
            Console.WriteLine("No revisions found.");
            Console.WriteLine();
            return;
        }

        foreach (Revision rev in revisions)
        {
            string revText = rev.ParentNode?.GetText()?.Trim() ?? "(no text)";
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            Console.WriteLine($"Affected Text: {revText}");
            Console.WriteLine();
        }
    }
}
