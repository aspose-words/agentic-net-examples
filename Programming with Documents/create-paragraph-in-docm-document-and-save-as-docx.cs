using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a blank document and save it as a macro‑enabled DOCM.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // creates a blank document
        string docmPath = "Sample.docm";
        doc.Save(docmPath, SaveFormat.Docm);              // save as DOCM (uses Save(string, SaveFormat))

        // ---------------------------------------------------------------
        // 2. Load the DOCM document, add a new paragraph, and save as DOCX.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docmPath);       // load the DOCM file
        DocumentBuilder builder = new DocumentBuilder(loadedDoc);
        builder.Writeln("This is a newly added paragraph."); // insert paragraph text

        string docxPath = "Sample.docx";
        loadedDoc.Save(docxPath);                         // save as DOCX (extension determines format)
    }
}
