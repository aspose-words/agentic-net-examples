using System;
using Aspose.Words;
using Aspose.Words.Settings;

namespace AsposeWordsPrerequisitesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Use DocumentBuilder to add some content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document demonstrates basic prerequisites for a DOC file.");

            // Set a built‑in document property (e.g., Title).
            doc.BuiltInDocumentProperties.Title = "Prerequisites Demo";

            // Optimize the document for a specific Word version (Word 2007 in this case).
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2007);

            // Save the document to the file system. The format is inferred from the extension.
            doc.Save("Prerequisites.docx");
        }
    }
}
