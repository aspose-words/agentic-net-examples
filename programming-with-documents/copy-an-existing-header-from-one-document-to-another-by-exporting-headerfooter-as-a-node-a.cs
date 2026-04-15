using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // 1. Create a source document that contains a header.
        // -----------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Move the cursor into the primary header and write some text.
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Write("Source Header Text");

        // Add a body paragraph so the document is not empty.
        srcBuilder.MoveToDocumentEnd();
        srcBuilder.Writeln("Source body paragraph.");

        // Optional: save the source file (helps with debugging).
        srcDoc.Save("Source.docx");

        // -----------------------------
        // 2. Create a destination (target) document.
        // -----------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination body paragraph.");

        // -----------------------------
        // 3. Export the header node from the source document.
        // -----------------------------
        HeaderFooter srcHeader = srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // -----------------------------
        // 4. Import the header node into the destination document.
        // -----------------------------
        Node importedHeader = dstDoc.ImportNode(srcHeader, true);

        // Add the imported header to the destination's first section.
        dstDoc.FirstSection.HeadersFooters.Add((HeaderFooter)importedHeader);

        // -----------------------------
        // 5. Save the resulting document.
        // -----------------------------
        dstDoc.Save("Destination.docx");
    }
}
