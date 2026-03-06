using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SaveChaptersAsHtml
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\MyDocument.docx";

        // Folder where the split HTML files will be written.
        string outputFolder = @"C:\Docs\Output\";
        Directory.CreateDirectory(outputFolder);

        // Folder for images that are extracted during HTML conversion.
        string imagesFolder = Path.Combine(outputFolder, "Images");
        Directory.CreateDirectory(imagesFolder);

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Configure HTML save options to split the document at each Heading 1 paragraph
        // (commonly used for chapters).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1, // split only at Heading 1 styles
            ImagesFolder = imagesFolder      // store linked images in a sub‑folder
        };

        // Save the document. Aspose.Words will create one HTML file per chapter.
        // The first file gets the name supplied below; subsequent parts receive
        // automatically generated names (e.g., MyDocument-01.html, MyDocument-02.html, …).
        string firstHtmlFile = Path.Combine(outputFolder, "MyDocument.html");
        doc.Save(firstHtmlFile, saveOptions);
    }
}
