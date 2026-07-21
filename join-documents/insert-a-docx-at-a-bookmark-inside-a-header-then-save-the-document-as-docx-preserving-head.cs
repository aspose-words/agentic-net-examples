using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertDocxIntoHeaderBookmark
{
    public static void Main()
    {
        // Define file names in the current directory.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        string sourceDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");

        // -------------------------------------------------
        // 1. Create the source document that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Font.Name = "Arial";
        srcBuilder.Font.Size = 14;
        srcBuilder.Font.Color = System.Drawing.Color.DarkBlue;
        srcBuilder.Writeln("This is the inserted document content.");
        // Save the source document (required by the rule to have a physical file).
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the destination document with a header containing a bookmark.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Ensure the document has a header.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        // Write some text before the bookmark.
        destBuilder.Writeln("Header start - ");

        // Insert a bookmark where the source document will be placed.
        string bookmarkName = "InsertHere";
        destBuilder.StartBookmark(bookmarkName);
        // Placeholder text inside the bookmark (can be empty).
        destBuilder.Writeln("[Placeholder]");
        destBuilder.EndBookmark(bookmarkName);

        // Write some text after the bookmark.
        destBuilder.Writeln(" - Header end.");

        // -------------------------------------------------
        // 3. Insert the source document at the bookmark inside the header.
        // -------------------------------------------------
        // Move the cursor back to the header and to the bookmark.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        destBuilder.MoveToBookmark(bookmarkName);
        // Insert the source document, preserving its formatting.
        destBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 4. Save the resulting document.
        // -------------------------------------------------
        destDoc.Save(outputDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 5. Validation: ensure the file exists and contains the inserted text.
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        Document validationDoc = new Document(outputDocPath);
        string fullText = validationDoc.GetText();

        if (!fullText.Contains("This is the inserted document content."))
            throw new InvalidOperationException("The inserted content was not found in the output document.");

        // Cleanup temporary source file (optional).
        if (File.Exists(sourceDocPath))
            File.Delete(sourceDocPath);
    }
}
