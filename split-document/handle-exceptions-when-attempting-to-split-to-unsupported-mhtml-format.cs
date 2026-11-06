using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections (two pages).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 - Page 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Page 2");

        // Save the source document (optional, just to have a file on disk).
        const string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Attempt to split the document by extracting the first page
        // and then saving the extracted part as MHTML, which is not supported.
        try
        {
            // Extract the first page into a new Document instance.
            Document pageDoc = doc.ExtractPages(1, 1);

            // Try to save the extracted page as MHTML.
            const string mhtmlPath = "Page1.mhtml";
            pageDoc.Save(mhtmlPath, SaveFormat.Mhtml);

            // If no exception is thrown, indicate unexpected success.
            Console.WriteLine("Unexpectedly saved the split page as MHTML.");
        }
        catch (Exception ex)
        {
            // Expected path: handle the exception for unsupported MHTML format.
            Console.WriteLine($"Caught exception while saving split document as MHTML: {ex.Message}");
        }
    }
}
