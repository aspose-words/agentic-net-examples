using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents are free of revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Configure comparison options (customize as needed).
        CompareOptions options = new CompareOptions
        {
            // Example: track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Do not ignore formatting; set to true to ignore formatting changes.
            IgnoreFormatting = false,
            // Use the edited document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, options);

        // Build a Markdown report describing the revisions.
        StringBuilder md = new StringBuilder();
        md.AppendLine("# Document Comparison Report");
        md.AppendLine();
        md.AppendLine($"Compared **Original.docx** with **Edited.docx** on {DateTime.Now}.");
        md.AppendLine();

        if (docOriginal.Revisions.Count == 0)
        {
            md.AppendLine("No differences were found.");
        }
        else
        {
            md.AppendLine($"Found **{docOriginal.Revisions.Count}** revisions:");
            md.AppendLine();

            int index = 1;
            foreach (Revision rev in docOriginal.Revisions)
            {
                md.AppendLine($"## Revision {index}");
                md.AppendLine();
                md.AppendLine($"- **Type:** {rev.RevisionType}");
                md.AppendLine($"- **Location:** {rev.ParentNode.NodeType}");
                md.AppendLine($"- **Changed Text:** `{rev.ParentNode.GetText().Trim()}`");
                md.AppendLine();
                index++;
            }
        }

        // Save the Markdown report to a file.
        System.IO.File.WriteAllText("ComparisonReport.md", md.ToString());

        // Optional: accept all revisions to make the original document identical to the edited one.
        // docOriginal.Revisions.AcceptAll();
        // docOriginal.Save("Merged.docx");
    }
}
