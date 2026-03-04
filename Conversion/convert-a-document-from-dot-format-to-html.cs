using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOT (Word template) file.
        Document doc = new Document("Template.dot");

        // Save the document in HTML format.
        doc.Save("Output.html", SaveFormat.Html);
    }
}
