using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for all files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonDemo");
        Directory.CreateDirectory(workDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder origBuilder = new DocumentBuilder(original);
        origBuilder.Writeln("This is a sample paragraph.");
        origBuilder.Font.Bold = true;
        origBuilder.Writeln("Bold text line.");
        string originalPath = Path.Combine(workDir, "Original.docx");
        original.Save(originalPath, SaveFormat.Docx);

        // Create the revised document with formatting changes and content changes.
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);
        // Change formatting: remove bold from the second line.
        revBuilder.Font.Bold = false;
        revBuilder.Writeln("Bold text line."); // same text, different formatting.
        // Content change: add a new paragraph.
        revBuilder.Writeln("An inserted new paragraph.");
        // Content change: modify first paragraph text.
        revBuilder.MoveToDocumentStart();
        revBuilder.Writeln("This is a modified sample paragraph.");
        string revisedPath = Path.Combine(workDir, "Revised.docx");
        revised.Save(revisedPath, SaveFormat.Docx);

        // Load documents (demonstrating load from file).
        Document baseDoc = new Document(originalPath);
        Document changedDoc = new Document(revisedPath);

        // Perform comparison. Revisions will be added to baseDoc.
        baseDoc.Compare(changedDoc, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (baseDoc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // Accept only formatting revisions, reject all others.
        // Copy revisions to a list to avoid collection modification issues.
        List<Revision> revisionList = new List<Revision>();
        foreach (Revision rev in baseDoc.Revisions)
        {
            revisionList.Add(rev);
        }

        foreach (Revision rev in revisionList)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();
            else
                rev.Reject();
        }

        // After processing, there should be no remaining revisions.
        if (baseDoc.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should have been processed, but some remain.");
        }

        // Save the final document.
        string resultPath = Path.Combine(workDir, "Result.docx");
        baseDoc.Save(resultPath, SaveFormat.Docx);
    }
}
