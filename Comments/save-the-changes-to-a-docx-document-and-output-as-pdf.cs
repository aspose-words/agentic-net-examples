using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Optionally rebuild the page layout before converting to PDF.
            // This ensures that any modifications made after loading are reflected.
            doc.UpdatePageLayout();

            // Save the document as PDF. The format is inferred from the file extension.
            doc.Save("OutputDocument.pdf");
        }
    }
}
