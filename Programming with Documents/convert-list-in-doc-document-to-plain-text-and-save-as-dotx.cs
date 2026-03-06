using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextTemplate
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourceDocPath = "InputDocument.doc";

        // Path where the resulting DOTX template will be saved.
        string outputTemplatePath = "ListPlainTextTemplate.dotx";

        // Load the existing DOC document.
        Document sourceDocument = new Document(sourceDocPath);

        // Configure text save options to simplify list labels for plain‑text conversion.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true   // Convert complex list symbols to simple ASCII equivalents.
        };

        // Export the document content to a plain‑text string using the configured options.
        string plainText = sourceDocument.ToString(txtOptions);

        // Create a new blank document that will become the DOTX template.
        Document templateDocument = new Document();

        // Insert the extracted plain‑text list into the new document.
        DocumentBuilder builder = new DocumentBuilder(templateDocument);
        builder.Writeln(plainText);

        // Save the document as a DOTX template. The format is inferred from the file extension.
        templateDocument.Save(outputTemplatePath);
    }
}
