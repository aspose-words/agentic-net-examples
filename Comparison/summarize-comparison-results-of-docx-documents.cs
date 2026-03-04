using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.AI;

class DocumentComparisonSummarizer
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";

        // Load the documents.
        Document originalDoc = new Document(originalPath);
        Document editedDoc   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (originalDoc.Revisions.Count != 0 || editedDoc.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

        // Gather revision details into a plain‑text summary.
        StringBuilder revisionsInfo = new StringBuilder();
        revisionsInfo.AppendLine("Document comparison results:");
        foreach (Revision rev in originalDoc.Revisions)
        {
            revisionsInfo.AppendLine($"- Type: {rev.RevisionType}");
            revisionsInfo.AppendLine($"  Node: {rev.ParentNode.NodeType}");
            revisionsInfo.AppendLine($"  Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Create a temporary document that contains the revision information.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Writeln(revisionsInfo.ToString());

        // Prepare the AI model for summarization.
        string apiKey = Environment.GetEnvironmentVariable("API_KEY");
        AiModel model = ((OpenAiModel)AiModel.Create(AiModelType.Gpt4OMini)
                                            .WithApiKey(apiKey))
                                            .WithOrganization("MyOrg")
                                            .WithProject("MyProject");

        // Set summarization options (e.g., short summary).
        SummarizeOptions summarizeOptions = new SummarizeOptions
        {
            SummaryLength = SummaryLength.Short
        };

        // Generate the summary document.
        Document summaryDoc = model.Summarize(tempDoc, summarizeOptions);

        // Save the summary.
        string summaryPath = @"C:\Docs\ComparisonSummary.docx";
        summaryDoc.Save(summaryPath);
    }
}
