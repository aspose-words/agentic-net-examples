using System;
using Aspose.Words;
using Aspose.Words.Saving;

class AppendFootnotesAndConvertToPdf
{
    static void Main()
    {
        // Create the destination document.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("This is the destination document.");
        dstBuilder.Writeln("[Destination footnote 1]");

        // Create the source document.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document.");
        srcBuilder.Writeln("[Source footnote 1]");
        srcBuilder.Writeln("[Source footnote 2]");

        // Append the source document to the destination document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        dstDoc.Save("CombinedResult.pdf", SaveFormat.Pdf);
    }
}
