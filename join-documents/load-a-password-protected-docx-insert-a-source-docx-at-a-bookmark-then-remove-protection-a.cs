using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;      // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary and output files
        string protectedDocPath = "protected.docx";
        string sourceDocPath = "source.docx";
        string outputPath = "merged.docx";

        // -----------------------------------------------------------------
        // Create a source DOCX that will be inserted later
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document content.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a password‑protected DOCX with a bookmark named "InsertHere"
        // -----------------------------------------------------------------
        Document protectedDoc = new Document();
        DocumentBuilder protBuilder = new DocumentBuilder(protectedDoc);
        protBuilder.Writeln("This is the protected document before insertion.");
        protBuilder.StartBookmark("InsertHere");
        protBuilder.Writeln("Bookmark location.");
        protBuilder.EndBookmark("InsertHere");

        // Apply read‑only protection and set a password for the file itself
        protectedDoc.Protect(ProtectionType.ReadOnly, "pwd123");
        OoxmlSaveOptions protectSaveOptions = new OoxmlSaveOptions { Password = "pwd123" };
        protectedDoc.Save(protectedDocPath, protectSaveOptions);

        // -----------------------------------------------------------------
        // Load the protected document using the password
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions { Password = "pwd123" };
        Document destDoc = new Document(protectedDocPath, loadOptions);

        // Load the source document that will be inserted
        Document srcDoc = new Document(sourceDocPath);

        // -----------------------------------------------------------------
        // Insert the source document at the bookmark location
        // -----------------------------------------------------------------
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.MoveToBookmark("InsertHere");
        destBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Remove protection and save the merged document
        // -----------------------------------------------------------------
        destDoc.Unprotect("pwd123");
        destDoc.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains the source text
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document mergedDoc = new Document(outputPath);
        string mergedText = mergedDoc.GetText();
        if (!mergedText.Contains("This is the source document content."))
            throw new InvalidOperationException("Source content was not inserted into the merged document.");
    }
}
