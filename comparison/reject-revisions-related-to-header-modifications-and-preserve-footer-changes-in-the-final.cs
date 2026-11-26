using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create the original document with a header, footer and body.
        // ------------------------------------------------------------
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
        builder.Writeln("Body content.");

        // ------------------------------------------------------------
        // Create the revised document with modified header, footer and body.
        // ------------------------------------------------------------
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        // Modified Header
        revBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        revBuilder.Writeln("Modified Header");

        // Modified Footer
        revBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        revBuilder.Writeln("Modified Footer");

        // Modified Body
        revBuilder.MoveToDocumentEnd();
        revBuilder.Writeln("Body content changed.");

        // ------------------------------------------------------------
        // Compare the documents. Revisions are added to the original.
        // ------------------------------------------------------------
        original.Compare(revised, "Comparer", DateTime.Now);

        // ------------------------------------------------------------
        // Process revisions:
        //   - Reject any changes that belong to a header.
        //   - Accept all other changes (including footer modifications).
        // ------------------------------------------------------------
        var revisionSnapshot = original.Revisions.Cast<Revision>().ToList(); // Snapshot to avoid collection modification issues.

        foreach (Revision rev in revisionSnapshot)
        {
            // Walk up the node hierarchy to find a HeaderFooter ancestor, if any.
            Node current = rev.ParentNode;
            HeaderFooter? headerFooter = null;

            while (current != null && !(current is HeaderFooter))
            {
                current = current.ParentNode;
            }

            if (current is HeaderFooter hf)
                headerFooter = hf;

            if (headerFooter != null)
            {
                // Determine whether the revision is inside a header.
                bool isHeader = headerFooter.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                                headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                                headerFooter.HeaderFooterType == HeaderFooterType.HeaderEven;

                if (isHeader)
                {
                    rev.Reject(); // Discard header modifications.
                }
                else
                {
                    rev.Accept(); // Keep footer modifications.
                }
            }
            else
            {
                // For revisions outside headers/footers (e.g., body), accept them.
                rev.Accept();
            }
        }

        // ------------------------------------------------------------
        // Save the final document where header changes are rejected
        // and footer changes are preserved.
        // ------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        original.Save(outputPath, SaveFormat.Docx);

        // Optional console output to verify revision count after processing.
        Console.WriteLine($"Final revision count: {original.Revisions.Count}");
        Console.WriteLine($"Result document saved to: {outputPath}");
    }
}
