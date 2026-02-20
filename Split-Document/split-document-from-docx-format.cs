using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentExample
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Save the document; each section will be written to a separate HTML file.
        doc.Save("OutputDocument.html", saveOptions);
    }
}
