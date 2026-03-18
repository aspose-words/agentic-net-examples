using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class ExtractAndEncrypt
{
    static void Main()
    {
        // Input and output file paths.
        const string inputPath = @"C:\Docs\Input.docx";
        const string outputPath = @"C:\Docs\EncryptedOutput.docx";

        // XPath expressions that locate the start and end nodes.
        // Adjust these expressions to match the actual nodes in your document.
        const string startXPath = "//w:p[w:r/w:t='START']";
        const string endXPath   = "//w:p[w:r/w:t='END']";

        // Load the source document.
        Document sourceDoc = new Document(inputPath);

        // Locate the start and end nodes.
        Node startNode = sourceDoc.SelectSingleNode(startXPath);
        Node endNode   = sourceDoc.SelectSingleNode(endXPath);

        if (startNode == null || endNode == null)
                throw new InvalidOperationException("Start or end node not found.");

        // Create a new empty document that will hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.EnsureMinimum(); // Guarantees a section, body, and paragraph exist.

        // Use NodeImporter for efficient node import while preserving formatting.
        NodeImporter importer = new NodeImporter(
            sourceDoc,               // Document to import from.
            extractedDoc,            // Document to import into.
            ImportFormatMode.KeepSourceFormatting);

        // Walk the node tree from the start node to the end node (inclusive)
        // and import each node into the new document.
        Node current = startNode;
        while (current != null)
        {
            Node imported = importer.ImportNode(current, true);
            // Append the imported node to the body of the new document.
            extractedDoc.FirstSection.Body.AppendChild(imported);

            if (current == endNode)
                break;

            // Move to the next node in a pre‑order traversal.
            current = current.NextPreOrder(sourceDoc);
        }

        // Prepare save options with a password to encrypt the document.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "MySecretPassword";

        // Save the extracted document using the password‑protected options.
        extractedDoc.Save(outputPath, saveOptions);
    }
}
