using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        Document doc = new Document("Input.doc");

        // Convert and save the document as PDF.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
