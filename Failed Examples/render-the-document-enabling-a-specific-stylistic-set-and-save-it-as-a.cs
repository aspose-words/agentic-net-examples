// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Enable a specific stylistic set (e.g., Stylistic Set 1) for every run in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            foreach (Run run in paragraph.GetChildNodes(NodeType.Run, true))
            {
                run.Font.StylisticSet = 1; // activates Stylistic Set 1
            }
        }

        // Create PDF save options (default settings can be customized if needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the modified document as a PDF file.
        doc.Save("Output.pdf", pdfOptions);
    }
}
