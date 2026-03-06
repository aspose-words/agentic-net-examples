using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.AI; // Namespace for AI models and related classes

namespace DocumentSplitAndSummarize
{
    // Callback that collects each split part as a Document object.
    public class CollectDocumentPartsCallback : IDocumentPartSavingCallback
    {
        private readonly List<Document> _parts;

        public CollectDocumentPartsCallback(List<Document> parts)
        {
            _parts = parts;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // The split part is provided as a Node (usually a Section).
            // Create a new Document and import the node so we can work with a full Document later.
            Document partDoc = new Document();
            Node importedNode = partDoc.ImportNode(args.Document, true);
            partDoc.AppendChild(importedNode);

            // Store the newly created Document.
            _parts.Add(partDoc);

            // Prevent actual file creation by assigning a dummy file name.
            args.DocumentPartFileName = "dummy.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the source DOCX document.
            Document sourceDoc = new Document("InputDocument.docx");

            // 2. Prepare a list to hold the split parts.
            List<Document> splitParts = new List<Document>();

            // 3. Configure HtmlSaveOptions to split the document by section breaks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new CollectDocumentPartsCallback(splitParts)
            };

            // 4. Save the document to a memory stream to trigger the splitting logic.
            using (MemoryStream dummyStream = new MemoryStream())
            {
                sourceDoc.Save(dummyStream, saveOptions);
            }

            // 5. Initialize the AI model (example uses OpenAI's GPT-4o mini).
            // Replace "YOUR_API_KEY" with a valid API key.
            var aiModel = AiModel.Create(AiModelType.Gpt4OMini)
                               .WithApiKey("YOUR_API_KEY");

            // 6. Summarize each split part and save the summaries.
            int partIndex = 1;
            foreach (Document part in splitParts)
            {
                // Configure summarization options (short summary in this example).
                SummarizeOptions summarizeOptions = new SummarizeOptions
                {
                    SummaryLength = SummaryLength.Short
                };

                // Generate the summary using the AI model.
                Document summaryDoc = aiModel.Summarize(part, summarizeOptions);

                // Save the summary to a separate DOCX file.
                string summaryFileName = $"Summary_Part_{partIndex}.docx";
                summaryDoc.Save(summaryFileName);
                partIndex++;
            }
        }
    }
}
