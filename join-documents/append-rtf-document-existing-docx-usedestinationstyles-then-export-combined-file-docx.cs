using System;
using Aspose.Words;

class AppendRtfToDocx
{
    static void Main()
    {
        // Path to the existing DOCX file (destination document).
        string destinationDocxPath = "Destination.docx";

        // Path to the source RTF file that will be appended.
        string sourceRtfPath = "Source.rtf";

        // Path where the combined DOCX will be saved.
        string outputDocxPath = "Combined.docx";

        // Load the destination DOCX document.
        Document dstDoc = new Document(destinationDocxPath);

        // Load the source RTF document.
        Document srcDoc = new Document(sourceRtfPath);

        // Append the RTF document to the end of the DOCX using destination styles.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as DOCX.
        dstDoc.Save(outputDocxPath);
    }
}
