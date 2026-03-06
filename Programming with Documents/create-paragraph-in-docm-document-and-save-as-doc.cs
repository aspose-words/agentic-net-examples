using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths for temporary files
        string docmPath = Path.Combine(Environment.CurrentDirectory, "TempDocument.docm");
        string docPath = Path.Combine(Environment.CurrentDirectory, "ResultDocument.doc");

        // -------------------------------------------------
        // 1. Create a blank document and save it as DOCM
        // -------------------------------------------------
        Document blankDoc = new Document(); // create a blank document

        // Use OoxmlSaveOptions to specify DOCM format
        OoxmlSaveOptions docmSaveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        blankDoc.Save(docmPath, docmSaveOptions); // save as DOCM

        // -------------------------------------------------
        // 2. Load the DOCM document
        // -------------------------------------------------
        Document docmDocument = new Document(docmPath); // load the DOCM file

        // -------------------------------------------------
        // 3. Insert a new paragraph using DocumentBuilder
        // -------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(docmDocument);
        builder.Writeln("This is a new paragraph added to the DOCM document.");

        // -------------------------------------------------
        // 4. Save the modified document as DOC (Word 97-2007)
        // -------------------------------------------------
        docmDocument.Save(docPath, SaveFormat.Doc); // save as DOC

        // Optional: clean up temporary DOCM file
        if (File.Exists(docmPath))
        {
            File.Delete(docmPath);
        }

        Console.WriteLine("Document saved as DOC at: " + docPath);
    }
}
