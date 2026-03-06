using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the macro-enabled DOCM document from disk.
        Document doc = new Document("Input.docm");

        // Save the document as MHTML (Web archive) using the explicit SaveFormat.
        doc.Save("Output.mhtml", SaveFormat.Mhtml);
    }
}
