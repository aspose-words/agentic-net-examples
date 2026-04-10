using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original document ----------
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Body content.
        builder.Writeln("Original body paragraph.");

        // Header content.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        // Footer content.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        // ---------- Create the edited document with changes ----------
        Document docEdited = (Document)docOriginal.Clone(true);
        DocumentBuilder editBuilder = new DocumentBuilder(docEdited);

        // Change body text.
        editBuilder.MoveToDocumentStart();
        editBuilder.Writeln("Edited body paragraph.");

        // Change header text.
        editBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        editBuilder.Writeln("Edited Header");

        // Change footer text.
        editBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        editBuilder.Writeln("Edited Footer");

        // ---------- Compare documents to generate revisions ----------
        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // ---------- Process revisions ----------
        // Reject revisions that belong to the header, accept all others.
        // Use a snapshot of the revisions to avoid collection modification issues.
        Revision[] revisions = docOriginal.Revisions.ToArray();
        foreach (Revision rev in revisions)
        {
            // Determine if the revision is inside a header.
            HeaderFooter? hf = rev.ParentNode?.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (hf != null && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
            {
                rev.Reject(); // Discard header changes.
            }
            else
            {
                rev.Accept(); // Keep all other changes (including footer).
            }
        }

        // ---------- Save the final document ----------
        string resultPath = Path.Combine(outputDir, "Result.docx");
        docOriginal.Save(resultPath);
    }
}
