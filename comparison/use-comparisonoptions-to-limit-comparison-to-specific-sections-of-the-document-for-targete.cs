using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Body content.
        builder.Writeln("Section 1 - Introduction.");

        // Header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        // Footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        // Return to the body to add more elements.
        builder.MoveTo(original.FirstSection.Body.LastParagraph);

        // Table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Comment.
        Comment comment = new Comment(original, "John Doe", "JD", DateTime.Now);
        comment.SetText("Original comment.");
        builder.CurrentParagraph.AppendChild(comment);

        // Save the original document (optional, for inspection).
        original.Save("Original.docx");

        // Clone the original and make targeted changes.
        Document revised = (Document)original.Clone(true);

        // Update body text.
        Paragraph bodyParagraph = revised.FirstSection.Body.FirstParagraph;
        if (bodyParagraph.Runs.Count > 0)
            bodyParagraph.Runs[0].Text = "Section 1 - Updated introduction.";

        // Update header text.
        HeaderFooter revisedHeader = revised.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (revisedHeader.FirstParagraph?.Runs.Count > 0)
            revisedHeader.FirstParagraph.Runs[0].Text = "Revised Header";

        // Save the revised document (optional, for inspection).
        revised.Save("Revised.docx");

        // Configure compare options to ignore headers, footers, tables, and comments.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreHeadersAndFooters = true,
            IgnoreTables = true,
            IgnoreComments = true,
            Target = ComparisonTargetType.New // Compare against the revised document as the target.
        };

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        original.Save("ComparisonResult.docx");

        // Output a summary of revisions that were detected (should reflect only body changes).
        Console.WriteLine($"Total revisions after applying CompareOptions: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            string parentInfo = rev.ParentNode?.ParentNode?.NodeType.ToString() ?? "N/A";
            string text = rev.ParentNode?.GetText().Trim() ?? string.Empty;
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Parent Node Type: {parentInfo}, Text: \"{text}\"");
        }
    }
}
