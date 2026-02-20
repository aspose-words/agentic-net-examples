using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // (Optional) Add some content to demonstrate that the document is not empty.
        // DocumentBuilder can be used to insert text, but it's not required for the conversion.
        // var builder = new DocumentBuilder(doc);
        // builder.Writeln("Sample text in DOC format.");

        // Specify save options for the DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document to a .doc file using the specified options.
        doc.Save("Result.doc", saveOptions);

        Console.WriteLine("Document created and saved as DOC format.");
    }
}
