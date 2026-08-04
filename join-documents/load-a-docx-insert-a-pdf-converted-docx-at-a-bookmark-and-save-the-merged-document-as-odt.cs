using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string baseDocPath = "BaseDocument.docx";
        string insertDocPath = "InsertedDocument.docx";
        string mergedDocPath = "MergedDocument.odt";

        // -----------------------------------------------------------------
        // 1. Create the base DOCX with a bookmark where the content will be inserted.
        // -----------------------------------------------------------------
        Document baseDoc = new Document();
        DocumentBuilder baseBuilder = new DocumentBuilder(baseDoc);
        baseBuilder.Writeln("This is the beginning of the base document.");
        baseBuilder.StartBookmark("InsertHere");
        baseBuilder.Writeln("[Placeholder for inserted content]");
        baseBuilder.EndBookmark("InsertHere");
        baseBuilder.Writeln("This is the end of the base document.");
        baseDoc.Save(baseDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a second DOCX that simulates a PDF‑to‑DOCX conversion.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("Content that originated from a PDF file.");
        insertBuilder.Writeln("Additional converted paragraph.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the base document.
        // -----------------------------------------------------------------
        Document loadedBase = new Document(baseDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document to be inserted.
        // -----------------------------------------------------------------
        Document loadedInsert = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // 5. Move the builder to the bookmark and insert the second document.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(loadedBase);
        builder.MoveToBookmark("InsertHere");
        // InsertDocument keeps the source formatting.
        builder.InsertDocument(loadedInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 6. Save the merged document as ODT.
        // -----------------------------------------------------------------
        OdtSaveOptions odtOptions = new OdtSaveOptions();
        loadedBase.Save(mergedDocPath, odtOptions);

        // -----------------------------------------------------------------
        // 7. Validation: ensure the file exists and contains text from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Merged ODT file was not created.");

        Document resultDoc = new Document(mergedDocPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("This is the beginning of the base document.") ||
            !resultText.Contains("Content that originated from a PDF file.") ||
            !resultText.Contains("This is the end of the base document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Program completed successfully.
    }
}
