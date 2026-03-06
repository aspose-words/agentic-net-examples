using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Add a simple paragraph to the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));

        // Save the document in the legacy DOC format.
        doc.Save("Result.doc", SaveFormat.Doc);
    }
}
