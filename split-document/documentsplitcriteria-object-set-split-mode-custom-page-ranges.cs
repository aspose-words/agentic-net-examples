using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path)
        Document doc = new Document("input.docx");

        // Create a DocumentSplitCriteria variable.
        // Combine flags to split the document at explicit page breaks and at heading paragraphs.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph;

        // Configure HtmlSaveOptions to use the custom split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Apply the custom split mode.
            DocumentSplitCriteria = splitCriteria,

            // Optional: define up to which heading level the split should occur.
            DocumentSplitHeadingLevel = 3
        };

        // Save the document. The output will be split into multiple HTML parts according to the criteria.
        doc.Save("output.html", saveOptions);
    }
}
