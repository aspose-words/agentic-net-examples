using System;
using Aspose.Words;
using Aspose.Words.AI; // GoogleAiModel

class DocumentSplitter
{
    static void Main()
    {
        // Load the source DOCX document. (Lifecycle rule: load)
        string sourcePath = @"C:\Docs\SourceDocument.docx";
        Document sourceDoc = new Document(sourcePath);

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Initialise the Google AI model. (Lifecycle rule: init)
        // Replace "YOUR_GOOGLE_API_KEY" with a valid Google AI API key.
        var aiModel = new GoogleAiModel("YOUR_GOOGLE_API_KEY");

        // Loop through each page, extract it, summarise it, and save the summary.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // Extract a single page. (Lifecycle rule: extract pages)
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Summarise the extracted page. The Summarize method returns a Document.
            Document summaryDoc = aiModel.Summarize(pageDoc);

            // Save the summary document. (Lifecycle rule: save)
            string summaryPath = $@"C:\Docs\Summaries\Page_{pageIndex}_Summary.docx";
            summaryDoc.Save(summaryPath);
        }

        Console.WriteLine("Document split and summarization completed.");
    }
}
