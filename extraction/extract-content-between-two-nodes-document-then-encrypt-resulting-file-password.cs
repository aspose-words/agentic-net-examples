using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractAndEncrypt
{
    static void Main()
    {
        // Create a sample document in memory with START and END markers.
        Document sourceDoc = new Document();
        sourceDoc.EnsureMinimum();

        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This paragraph is before the start marker.");
        builder.Writeln("START");
        builder.Writeln("Paragraph 1 inside the range.");
        builder.Writeln("Paragraph 2 inside the range.");
        builder.Writeln("END");
        builder.Writeln("This paragraph is after the end marker.");

        // Locate the start and end paragraphs by their text content.
        Paragraph startNode = null;
        Paragraph endNode = null;

        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            string txt = para.GetText().Trim();
            if (txt == "START")
                startNode = para;
            else if (txt == "END")
                endNode = para;
        }

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

        // Walk through the paragraphs from startNode to endNode (inclusive)
        // and import each paragraph into the new document.
        Paragraph current = startNode;
        while (current != null)
        {
            Paragraph imported = (Paragraph)importer.ImportNode(current, true);
            extractedDoc.FirstSection.Body.AppendChild(imported);

            if (current == endNode)
                break;

            current = current.NextSibling as Paragraph;
        }

        // Prepare save options with a password to encrypt the document.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "MySecretPassword"
        };

        // Determine a temporary output path.
        string outputPath = Path.Combine(Path.GetTempPath(), "EncryptedOutput.docx");

        // Save the extracted document using the password‑protected options.
        extractedDoc.Save(outputPath, saveOptions);

        Console.WriteLine($"Encrypted document saved to: {outputPath}");
    }
}
