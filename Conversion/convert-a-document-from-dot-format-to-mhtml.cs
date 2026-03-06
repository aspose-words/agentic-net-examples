using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document in DOT (template) format.
        Document doc = new Document("InputTemplate.dot");

        // Save the document as MHTML (Web archive) format.
        doc.Save("OutputDocument.mht", SaveFormat.Mhtml);
    }
}
