using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.AI;

class Program
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document edited = new Document("Edited.docx");

        // Documents must not contain revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. Revisions are added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Collect revision details into plain text.
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Comparison Summary:");
        foreach (Revision rev in original.Revisions)
        {
            sb.AppendLine($"- Revision Type: {rev.RevisionType}, Node Type: {rev.ParentNode.NodeType}");
            sb.AppendLine($"  Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Create a temporary document that holds the raw revision text.
        Document rawSummaryDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(rawSummaryDoc);
        builder.Writeln(sb.ToString());

        // Prepare the AI model for summarization.
        string apiKey = Environment.GetEnvironmentVariable("API_KEY");
        AiModel model = AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);

        // Configure summarization options (short summary).
        SummarizeOptions summarizeOptions = new SummarizeOptions
        {
            SummaryLength = SummaryLength.Short
        };

        // Generate a concise summary using the AI model.
        Document conciseSummary = model.Summarize(rawSummaryDoc, summarizeOptions);

        // Save the final summary document.
        conciseSummary.Save("ComparisonResultSummary.docx");
    }
}
