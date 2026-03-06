using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextToDotm
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourceDocPath = @"C:\Docs\SourceList.doc";

        // Load the source document.
        Document sourceDoc = new Document(sourceDocPath);

        // Ensure that list labels are up‑to‑date (optional, but useful if the document
        // was edited programmatically before this step).
        sourceDoc.UpdateListLabels();

        // Extract the plain‑text representation of the entire document,
        // which includes the list items rendered as simple text.
        PlainTextDocument plainText = new PlainTextDocument(sourceDocPath);
        string listAsText = plainText.Text;

        // Create a new (blank) document that will become the DOTM template.
        Document dotmDoc = new Document();

        // Insert the extracted plain‑text into the new document.
        DocumentBuilder builder = new DocumentBuilder(dotmDoc);
        builder.Writeln(listAsText);

        // Save the result as a macro‑enabled Word template (DOTM).
        string outputDotmPath = @"C:\Docs\ListPlainTextTemplate.dotm";
        dotmDoc.Save(outputDotmPath, SaveFormat.Dotm);
    }
}
