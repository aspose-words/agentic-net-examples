using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load a source document. Aspose.Words can load DOCX, HTML, MHTML, EPUB, etc.
        Document doc = new Document("Source.docx");

        // -----------------------------------------------------------------
        // Convert to DOCM (macro‑enabled Word) using OoxmlSaveOptions.
        // The factory method creates the correct SaveOptions subclass for DOCM.
        // -----------------------------------------------------------------
        SaveOptions docmOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docm);

        // The returned object is OoxmlSaveOptions; cast to access OOXML‑specific settings.
        if (docmOptions is OoxmlSaveOptions ooxml)
        {
            // Example configuration: enforce strict OOXML compliance and set a password.
            ooxml.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            ooxml.Password = "SecretPassword";
        }

        // Save the document as DOCM.
        doc.Save("Result.docm", docmOptions);

        // -----------------------------------------------------------------
        // Additional conversions using HtmlSaveOptions for HTML, MHTML, and EPUB.
        // -----------------------------------------------------------------

        // Convert to HTML.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportTextInputFormFieldAsText = true,
            Encoding = Encoding.UTF8,
            PrettyFormat = true
        };
        doc.Save("Result.html", htmlOptions);

        // Convert to MHTML.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true,
            Encoding = Encoding.UTF8
        };
        doc.Save("Result.mhtml", mhtmlOptions);

        // Convert to EPUB.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            ExportDocumentProperties = true,
            Encoding = Encoding.UTF8
        };
        doc.Save("Result.epub", epubOptions);
    }
}
