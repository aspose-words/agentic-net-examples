using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Replacing;

namespace AsposeWordsComparisonDemo
{
    // Criteria that matches revisions occurring in any header.
    public class HeaderRevisionCriteria : IRevisionCriteria
    {
        public bool IsMatch(Revision revision)
        {
            if (revision?.ParentNode == null)
                return false;

            // Header/footer nodes have NodeType.HeaderFooter.
            if (revision.ParentNode.NodeType != NodeType.HeaderFooter)
                return false;

            var headerFooter = (HeaderFooter)revision.ParentNode;
            // Reject only header revisions, keep footers.
            return headerFooter.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                   headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                   headerFooter.HeaderFooterType == HeaderFooterType.HeaderEven;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create the original document with header, footer and body text.
            var original = new Document();
            var builder = new DocumentBuilder(original);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Original Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Original Footer");
            builder.MoveToDocumentEnd();
            builder.Writeln("Original body content.");

            // Create the revised document with changed header, footer and body text.
            var revised = new Document();
            var revBuilder = new DocumentBuilder(revised);
            revBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            revBuilder.Writeln("Edited Header");
            revBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            revBuilder.Writeln("Edited Footer");
            revBuilder.MoveToDocumentEnd();
            revBuilder.Writeln("Edited body content.");

            // Compare the documents – revisions are created in the original document.
            original.Compare(revised, "Comparer", DateTime.Now);

            // Reject all revisions that belong to headers.
            original.Revisions.Reject(new HeaderRevisionCriteria());

            // Accept all remaining revisions (including footer changes).
            original.Revisions.AcceptAll();

            // Save the final document where header changes are rejected and footer changes are kept.
            original.Save("Result.docx");
        }
    }
}
