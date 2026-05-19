using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the documents.
        const string destinationPath = "Destination.docx";
        const string sourcePath = "Source.docx";

        // -------------------------------------------------
        // Create the source DOCX that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        // Add some formatted content to the source document.
        srcBuilder.Font.Name = "Arial";
        srcBuilder.Font.Size = 14;
        srcBuilder.Font.Color = System.Drawing.Color.Blue;
        srcBuilder.Writeln("Inserted content with its own formatting.");
        // Save the source document as DOCX.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create the destination document with a header that contains a bookmark.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Create a primary header for the first section.
        HeaderFooter header = new HeaderFooter(destDoc, HeaderFooterType.HeaderPrimary);
        destDoc.FirstSection.HeadersFooters.Add(header);

        // Build the header content.
        DocumentBuilder headerBuilder = new DocumentBuilder(destDoc);
        headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        headerBuilder.Write("Header start ");
        // Insert an empty bookmark where the source document will be placed.
        headerBuilder.StartBookmark("InsertHere");
        headerBuilder.EndBookmark("InsertHere");
        headerBuilder.Writeln(" Header end.");

        // -------------------------------------------------
        // Insert the source document at the bookmark inside the header.
        // -------------------------------------------------
        headerBuilder.MoveToBookmark("InsertHere");
        // KeepSourceFormatting preserves the formatting of the inserted document.
        headerBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // Save the resulting document as DOCX, preserving header formatting.
        // -------------------------------------------------
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(destinationPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");
    }
}
