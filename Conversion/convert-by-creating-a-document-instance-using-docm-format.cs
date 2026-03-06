using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello DOCM!");

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("Result.docm", SaveFormat.Docm);
    }
}
