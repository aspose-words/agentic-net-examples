using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.AI;

class ComparisonSummary
{
    static void Main()
    {
        // Load the original and edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents; revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Reviewer", DateTime.Now);
        }

        // Create a temporary document that will hold a textual description of the revisions.
        Document revisionsReport = new Document();
        DocumentBuilder builder = new DocumentBuilder(revisionsReport);
        builder.Writeln("Comparison Summary:");
        builder.Writeln();

        // Iterate through each revision and write its details to the report.
        foreach (Revision rev in docOriginal.Revisions)
        {
            builder.Writeln($"Revision Type: {rev.RevisionType}");
            builder.Writeln($"Node Type: {rev.ParentNode.NodeType}");
            builder.Writeln($"Changed Text: \"{rev.ParentNode.GetText().Trim()}\"");
            builder.Writeln();
        }

        // Prepare the AI model for summarization.
        string apiKey = Environment.GetEnvironmentVariable("API_KEY");
        // No need for the OpenAI specific namespace; the generic AiModel works for supported providers.
        AiModel model = AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);

        // Configure summarization options (short summary in this example).
        SummarizeOptions summarizeOptions = new SummarizeOptions
        {
            SummaryLength = SummaryLength.Short
        };

        // Generate a concise summary of the revisions report using the AI model.
        Document finalSummary = model.Summarize(revisionsReport, summarizeOptions);

        // Save the summarized comparison result.
        finalSummary.Save("ComparisonSummary.docx");
    }
}
