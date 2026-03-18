using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph.
        builder.InsertParagraph();

        // Insert a DATE field with the format "MMMM dd, yyyy".
        // The field code is inserted without the surrounding braces.
        builder.InsertField(@"DATE \@ ""MMMM dd, yyyy""");

        // Save the document to disk.
        doc.Save("CurrentDateParagraph.docx");
    }
}
