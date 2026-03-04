using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath = @"C:\Docs\Edited.docx";

        // Load the documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");
        }

        // Set up comparison options (customize as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: ignore case changes and formatting.
            IgnoreCaseChanges = true,
            IgnoreFormatting = true,
            // Show changes in the edited document (new document as target).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Build a Markdown report of the revisions.
        StringBuilder mdReport = new StringBuilder();
        mdReport.AppendLine("# Document Comparison Report");
        mdReport.AppendLine();
        mdReport.AppendLine($"**Compared on:** {DateTime.Now}");
        mdReport.AppendLine();
        mdReport.AppendLine("## Revisions");
        mdReport.AppendLine();

        foreach (Revision rev in docOriginal.Revisions)
        {
            // Get a short description of the revision type.
            string revType = rev.RevisionType.ToString();

            // Get the node type where the revision occurred.
            string nodeType = rev.ParentNode.NodeType.ToString();

            // Get the text of the changed node (trimmed for readability).
            string changedText = rev.ParentNode.GetText().Trim().Replace("\r", " ").Replace("\n", " ");

            // Append a markdown list item.
            mdReport.AppendLine($"- **{revType}** in **{nodeType}**: `{changedText}`");
        }

        // Save the markdown report.
        string mdPath = @"C:\Docs\ComparisonReport.md";
        File.WriteAllText(mdPath, mdReport.ToString());

        // Optionally, save the original document with revisions applied.
        string revisedDocPath = @"C:\Docs\Original_WithRevisions.docx";
        docOriginal.Save(revisedDocPath);
    }
}
