using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the macro-enabled Word document (DOCM) from disk.
        Document doc = new Document("Input.docm");

        // Save the document as HTML. The file extension determines the format,
        // but we also explicitly specify SaveFormat.Html for clarity.
        doc.Save("Output.html", SaveFormat.Html);
    }
}
