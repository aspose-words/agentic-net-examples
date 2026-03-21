using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the destination document and add initial content.
        // -----------------------------------------------------------------
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("Content before the inserted document.");

        // -----------------------------------------------------------------
        // 2. Create the source document that will be inserted.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("This is the inserted document content.");

        // -----------------------------------------------------------------
        // 3. Insert the source document at the current cursor position.
        // -----------------------------------------------------------------
        destBuilder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);
        destBuilder.Writeln("Content after the inserted document.");

        // -----------------------------------------------------------------
        // 4. Save the combined document as HTML.
        // -----------------------------------------------------------------
        const string outputHtmlPath = "Result.html";
        destination.Save(outputHtmlPath, SaveFormat.Html);
    }
}
