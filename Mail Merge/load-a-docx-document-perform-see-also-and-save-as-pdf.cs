using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Update all fields (including any SEE ALSO fields) so that they reflect the latest content.
        doc.UpdateFields();

        // Recalculate the layout to ensure pagination is correct before conversion.
        doc.UpdatePageLayout();

        // Save the document as a PDF file.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
