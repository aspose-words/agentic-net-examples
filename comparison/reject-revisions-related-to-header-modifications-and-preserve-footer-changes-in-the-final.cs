using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the original document with header, footer and body content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Header
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        // Footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        // Body
        builder.MoveToDocumentEnd();
        builder.Writeln("Original body text.");

        // Create the revised version by cloning and editing header, footer and body.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        // Edit header
        revBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        revBuilder.Writeln("Edited Header");

        // Edit footer
        revBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        revBuilder.Writeln("Edited Footer");

        // Edit body
        revBuilder.MoveToDocumentEnd();
        revBuilder.Writeln("Edited body text.");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Helper to determine whether a revision belongs to a header.
        bool IsHeaderRevision(Revision rev)
        {
            Node? parent = rev.ParentNode;
            if (parent == null) return false;
            Node? grandParent = parent.ParentNode;
            if (grandParent == null) return false;
            if (grandParent.NodeType == NodeType.HeaderFooter)
            {
                HeaderFooter hf = (HeaderFooter)grandParent;
                return hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                       hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                       hf.HeaderFooterType == HeaderFooterType.HeaderEven;
            }
            return false;
        }

        // Reject header revisions, accept all others (including footer changes).
        for (int i = original.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = original.Revisions[i];
            if (IsHeaderRevision(rev))
                rev.Reject();   // Discard header modifications.
            else
                rev.Accept();   // Keep other modifications (e.g., footer, body).
        }

        // Save the final document where header changes are rejected and footer changes are preserved.
        original.Save("Result.docx");
    }
}
