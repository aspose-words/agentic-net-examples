using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Loading;

class ContentControlInTxt
{
    static void Main()
    {
        // Path to the source TXT file.
        string txtPath = "input.txt";

        // Load the TXT document with default load options.
        // TxtLoadOptions allows configuring TXT-specific loading behavior.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document doc = new Document(txtPath, loadOptions);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document so the content control is appended.
        builder.MoveToDocumentEnd();

        // Create an inline plain‑text Structured Document Tag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleTag",   // Friendly name shown in the UI.
            Tag = "SampleTagId",   // Identifier used for programmatic access.
            LockContents = true   // Prevent the user from editing the contents.
        };

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Add placeholder text inside the content control.
        sdt.AppendChild(new Run(doc, "Enter text here"));

        // Save the modified document as a DOCX file.
        doc.Save("output.docx");
    }
}
