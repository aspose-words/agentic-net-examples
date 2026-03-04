using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocx
{
    static void Main()
    {
        // Input DOCX file.
        string inputPath = @"C:\Docs\Sample.docx";

        // Load the document from file.
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Convert to HTML.
        // The Save method automatically determines the format from the extension.
        // -----------------------------------------------------------------
        doc.Save(@"C:\Docs\Sample.html", SaveFormat.Html);

        // -----------------------------------------------------------------
        // Convert to MHTML.
        // Use HtmlSaveOptions to specify the target format and additional options.
        // -----------------------------------------------------------------
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            Encoding = new UTF8Encoding(false),   // UTF‑8 without BOM.
            ExportDocumentProperties = true       // Include document properties in the output.
        };
        doc.Save(@"C:\Docs\Sample.mhtml", mhtmlOptions);

        // -----------------------------------------------------------------
        // Convert to EPUB.
        // HtmlSaveOptions also handles EPUB conversion.
        // -----------------------------------------------------------------
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            Encoding = new UTF8Encoding(false),
            ExportDocumentProperties = true,
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph, // Split by headings.
            DocumentSplitHeadingLevel = 2                                   // Up to heading level 2.
        };
        doc.Save(@"C:\Docs\Sample.epub", epubOptions);
    }
}
