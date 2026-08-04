using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with header, footer and body text.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        builder.MoveToDocumentEnd();
        builder.Writeln("Original body text.");

        // Create the revised document with modified header, footer and body text.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);

        builder2.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder2.Writeln("Edited Header");

        builder2.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder2.Writeln("Edited Footer");

        builder2.MoveToDocumentEnd();
        builder2.Writeln("Edited body text.");

        // Compare the documents – revisions are created in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Reject only the revisions that belong to the primary header.
        original.Revisions.Reject(new HeaderRevisionCriteria());

        // Accept all remaining revisions (footer and body changes are kept).
        original.Revisions.AcceptAll();

        // Save the final document.
        original.Save("Result.docx");
    }

    // Criteria that matches revisions located in the primary header.
    private class HeaderRevisionCriteria : IRevisionCriteria
    {
        public bool IsMatch(Revision revision)
        {
            if (revision?.ParentNode == null)
                return false;

            if (revision.ParentNode.NodeType == NodeType.HeaderFooter)
            {
                HeaderFooter hf = (HeaderFooter)revision.ParentNode;
                return hf.HeaderFooterType == HeaderFooterType.HeaderPrimary;
            }

            return false;
        }
    }
}
