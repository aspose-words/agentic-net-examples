using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document (replace with your own path if needed)
        Document doc = new Document("Input.docx");

        // Create HtmlSaveOptions which will be used for saving.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the split criteria to split the output into separate pages.
        // Using PageBreak will create a separate HTML part for each page.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.PageBreak;

        // Save the document. Each page will be saved as a separate HTML file.
        doc.Save("Output.html", saveOptions);
    }
}
