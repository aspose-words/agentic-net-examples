using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Instantiate a blank Word document.
        Document doc = new Document();

        // Add a simple paragraph to demonstrate that the document contains content.
        doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Hello DOC format!"));

        // Save the document explicitly in the legacy DOC format.
        doc.Save("Output.doc", SaveFormat.Doc);
    }
}
