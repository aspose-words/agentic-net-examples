using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Compare the documents. Do not ignore any element types so that header revisions are generated.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Reject all revisions that belong to headers, preserving footers.
        original.Revisions.Reject(new HeaderRevisionCriteria());

        // Save the resulting document. Footer changes remain, header changes are removed.
        original.Save("Result.docx");
    }

    // Custom criteria that matches revisions located inside header sections.
    private class HeaderRevisionCriteria : IRevisionCriteria
    {
        public bool IsMatch(Revision revision)
        {
            // Revisions without a parent node cannot be in a header.
            if (revision.ParentNode == null)
                return false;

            // Find the nearest HeaderFooter ancestor of the revision's parent node.
            Node headerFooterNode = revision.ParentNode.GetAncestor(NodeType.HeaderFooter);
            if (headerFooterNode == null)
                return false;

            // Cast to HeaderFooter and check if it is a header (not a footer).
            HeaderFooter headerFooter = (HeaderFooter)headerFooterNode;
            return headerFooter.HeaderFooterType.ToString().Contains("Header");
        }
    }
}
