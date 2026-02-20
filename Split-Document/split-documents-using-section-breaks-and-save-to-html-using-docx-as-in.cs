using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentBySection
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Base name for the output HTML files.
        // The first part will be saved as this name,
        // additional parts will have suffixes added automatically.
        string outputPath = "output.html";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Set up HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

        // Save the document. Aspose.Words will create separate HTML files
        // for each section according to the split criteria.
        doc.Save(outputPath, saveOptions);
    }
}
