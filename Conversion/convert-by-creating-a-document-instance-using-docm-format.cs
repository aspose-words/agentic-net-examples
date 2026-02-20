using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Optionally add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, DOCM!");

        // Prepare save options for the DOCM (macro‑enabled) format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Save the document as a DOCM file.
        doc.Save("Result.docm", saveOptions);
    }
}
