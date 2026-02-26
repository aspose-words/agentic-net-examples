using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Paths for the intermediate DOCM file and the final DOCX file.
            const string docmPath = "IntermediateDocument.docm";
            const string docxPath = "ResultDocument.docx";

            // -------------------------------------------------
            // 1. Create a new blank document (DOCM will be created later).
            // -------------------------------------------------
            Document doc = new Document();                     // Document() constructor
            DocumentBuilder builder = new DocumentBuilder(doc); // DocumentBuilder for editing

            // -------------------------------------------------
            // 2. Insert a paragraph into the document.
            // -------------------------------------------------
            builder.Writeln("This is a paragraph inside a macro‑enabled document.");

            // -------------------------------------------------
            // 3. Save the document as a macro‑enabled DOCM file.
            // -------------------------------------------------
            OoxmlSaveOptions saveAsDocm = new OoxmlSaveOptions(SaveFormat.Docm); // SaveOptions with DOCM format
            doc.Save(docmPath, saveAsDocm);                                         // Save(string, SaveOptions)

            // -------------------------------------------------
            // 4. Load the DOCM file we just created.
            // -------------------------------------------------
            Document loadedDoc = new Document(docmPath); // Document(string) constructor

            // -------------------------------------------------
            // 5. Save the loaded document as a regular DOCX file.
            // -------------------------------------------------
            loadedDoc.Save(docxPath, SaveFormat.Docx); // Save(string, SaveFormat)
        }
    }
}
