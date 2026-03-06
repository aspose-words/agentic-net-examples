using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        builder.Writeln("Hello again!");

        // Enable rendering of paragraph marks (pilcrow symbols) in the output.
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Save the document. The format is inferred from the file extension.
        doc.Save("LayoutOptionsParagraphMarks.pdf");
    }
}
