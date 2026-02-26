using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Save the document. The split criteria causes Aspose.Words to generate
        // separate HTML files for each section (e.g., output.html, output-01.html, etc.).
        doc.Save("output.html", saveOptions);
    }
}
