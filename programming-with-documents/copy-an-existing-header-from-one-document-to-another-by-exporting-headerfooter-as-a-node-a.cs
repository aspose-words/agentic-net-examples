using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace HeaderCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create the source document with a primary header.
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            srcBuilder.Write("Source Header Text");
            srcBuilder.MoveToDocumentEnd();
            srcBuilder.Writeln("Source document body.");

            // Create the destination document.
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            dstBuilder.Writeln("Destination document body.");

            // Export the header node from the source document.
            HeaderFooter srcHeader = srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

            // Import the header node into the destination document.
            Node importedHeaderNode = dstDoc.ImportNode(srcHeader, true, ImportFormatMode.KeepSourceFormatting);
            HeaderFooter importedHeader = (HeaderFooter)importedHeaderNode;

            // Add the imported header to the destination document's first section.
            dstDoc.FirstSection.HeadersFooters.Add(importedHeader);

            // Save the resulting document.
            dstDoc.Save("DestinationWithCopiedHeader.docx");
        }
    }
}
