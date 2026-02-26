using System;
using System.Text;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Load the PDF file as an Aspose.Words document.
        Document sourceDoc = new Document("input.pdf");

        // Define the range of paragraphs to extract (zero‑based indices).
        // For example, extract content from the 3rd paragraph up to (but not including) the 6th.
        int startParagraphIndex = 2; // third paragraph
        int endParagraphIndex   = 5; // sixth paragraph (exclusive)

        // Validate the indices against the actual paragraph count.
        ParagraphCollection paragraphs = sourceDoc.FirstSection.Body.Paragraphs;
        if (startParagraphIndex < 0 ||
            endParagraphIndex > paragraphs.Count ||
            startParagraphIndex >= endParagraphIndex)
        {
            throw new ArgumentOutOfRangeException("Invalid paragraph range specified.");
        }

        // Gather the text of the selected paragraphs.
        StringBuilder extractedText = new StringBuilder();
        for (int i = startParagraphIndex; i < endParagraphIndex; i++)
        {
            Paragraph para = paragraphs[i];
            extractedText.Append(para.GetText()); // GetText includes the paragraph break.
        }

        // Output the extracted content to the console.
        Console.WriteLine(extractedText.ToString());

        // Optionally, save the extracted content as a new Word document.
        Document resultDoc = new Document(); // create a blank document
        // Ensure the document has the minimal required nodes.
        resultDoc.RemoveAllChildren();
        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);
        Body body = new Body(resultDoc);
        section.AppendChild(body);
        Paragraph newPara = new Paragraph(resultDoc);
        body.AppendChild(newPara);
        newPara.AppendChild(new Run(resultDoc, extractedText.ToString()));

        resultDoc.Save("ExtractedContent.docx");
    }
}
