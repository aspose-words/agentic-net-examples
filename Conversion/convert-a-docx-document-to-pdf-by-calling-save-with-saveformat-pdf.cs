using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document from disk.
        Document doc = new Document("input.docx");

        // Save the document as PDF. The SaveFormat.Pdf enum tells Aspose.Words to use PDF output.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
