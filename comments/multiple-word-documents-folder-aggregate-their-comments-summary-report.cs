using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.AI;

class AggregateCommentsAndSummarize
{
    static void Main()
    {
        // Folder containing the source Word documents.
        string sourceFolder = @"C:\Docs";

        // Create a blank document that will hold all extracted comments.
        Document aggregatedDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(aggregatedDoc);
        builder.Writeln("Aggregated Comments:");
        builder.Writeln();

        // Load each .docx file, extract its comments, and write them into the aggregated document.
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            Document srcDoc = new Document(filePath);
            NodeCollection commentNodes = srcDoc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in commentNodes)
            {
                string commentText = comment.Range.Text.Trim();
                builder.Writeln($"From {Path.GetFileName(filePath)}: {commentText}");
            }
        }

        // Add a placeholder for the summary that will be generated later.
        builder.Writeln();
        builder.Writeln("Summary:");
        builder.Writeln();

        // Save the document that contains the raw comments.
        string commentsPath = Path.Combine(sourceFolder, "AggregatedComments.docx");
        aggregatedDoc.Save(commentsPath);

        // ------------------------------------------------------------
        // Generate a concise summary of the aggregated comments using an AI model.
        // ------------------------------------------------------------

        string apiKey = Environment.GetEnvironmentVariable("API_KEY");
        Document summaryDoc;

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            // No API key – create a simple placeholder summary.
            summaryDoc = aggregatedDoc.Clone();
            DocumentBuilder sumBuilder = new DocumentBuilder(summaryDoc);
            sumBuilder.Writeln();
            sumBuilder.Writeln("[Summary could not be generated because an OpenAI API key was not provided.]");
        }
        else
        {
            // Create an OpenAI model instance and configure it with the API key.
            OpenAiModel aiModel = (OpenAiModel)AiModel.Create(AiModelType.Gpt4OMini).WithApiKey(apiKey);

            // Configure summarization options (short summary in this example).
            SummarizeOptions summarizeOptions = new SummarizeOptions
            {
                SummaryLength = SummaryLength.Short
            };

            // Generate the summary document.
            summaryDoc = aiModel.Summarize(aggregatedDoc, summarizeOptions);
        }

        // Save the summary report.
        string summaryPath = Path.Combine(sourceFolder, "CommentsSummary.docx");
        summaryDoc.Save(summaryPath);
    }
}
