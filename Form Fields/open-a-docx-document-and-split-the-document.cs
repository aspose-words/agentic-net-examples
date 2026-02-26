using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Open the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Set up save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

        // Save the document. Aspose.Words will create multiple HTML files,
        // one for each section (e.g., Output.html, Output_1.html, ...).
        doc.Save("Output.html", saveOptions);
    }
}
