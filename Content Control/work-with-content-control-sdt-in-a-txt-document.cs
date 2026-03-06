using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Prepare a simple TXT content.
        string txtContent = "Hello world!\nThis is a sample text file.";

        // Load the TXT content into an Aspose.Words Document using TxtLoadOptions.
        using (MemoryStream txtStream = new MemoryStream())
        {
            byte[] txtBytes = System.Text.Encoding.UTF8.GetBytes(txtContent);
            txtStream.Write(txtBytes, 0, txtBytes.Length);
            txtStream.Position = 0; // Reset stream position for reading.

            TxtLoadOptions loadOptions = new TxtLoadOptions(); // Default options.
            Document doc = new Document(txtStream, loadOptions); // Load document.

            // Create a plain‑text Structured Document Tag (content control) at inline level.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "SampleTag",          // Friendly name.
                Tag = "SampleTagId",          // Tag identifier.
                IsShowingPlaceholderText = true // Show placeholder when empty.
            };

            // Insert the SDT at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.InsertNode(sdt);

            // Save the resulting document as DOCX.
            doc.Save("Output.docx");
        }
    }
}
