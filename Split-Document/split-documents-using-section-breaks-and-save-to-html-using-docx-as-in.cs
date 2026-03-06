using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Base path for the output HTML files.
        // The first part will be saved as this name,
        // subsequent parts will have suffixes like "-01.html", "-02.html", etc.
        string outputPath = "output.html";

        // Load the document from the DOCX file.
        Document doc = new Document(inputPath);

        // Set up HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Save the document. Aspose.Words will create multiple HTML files
        // according to the number of sections in the source document.
        doc.Save(outputPath, saveOptions);
    }
}
