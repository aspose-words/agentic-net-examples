using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class AppendRtfToDocx
{
    static void Main()
    {
        // Create temporary files for the example.
        string destinationDocxPath = Path.Combine(Path.GetTempPath(), "Destination.docx");
        string sourceRtfPath = Path.Combine(Path.GetTempPath(), "Source.rtf");
        string outputDocxPath = Path.Combine(Path.GetTempPath(), "Combined.docx");

        // Create a simple DOCX destination document.
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the content of the destination DOCX document.");
        destinationDoc.Save(destinationDocxPath, SaveFormat.Docx);

        // Create a simple RTF source document.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the source RTF document.");
        sourceDoc.Save(sourceRtfPath, SaveFormat.Rtf);

        // Load the documents from the temporary files.
        Document dstDoc = new Document(destinationDocxPath);
        Document srcRtfDoc = new Document(sourceRtfPath);

        // Append the RTF document to the end of the DOCX using destination styles.
        dstDoc.AppendDocument(srcRtfDoc, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        dstDoc.Save(outputDocxPath, SaveFormat.Docx);

        Console.WriteLine("Combined document saved to: " + outputDocxPath);
    }
}
