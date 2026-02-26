using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a couple of paragraphs.
        builder.Writeln("Hello world!");
        builder.Writeln("Hello again!");

        // Enable rendering of paragraph marks (pilcrow symbols) in the output.
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Save the document. PDF format will display the paragraph marks.
        doc.Save("LayoutOptionsParagraphMarks.pdf");
    }
}
