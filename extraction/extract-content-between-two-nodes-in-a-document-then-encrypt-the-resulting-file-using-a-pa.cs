using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample source document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Intro paragraph.");
        builder.Writeln("Start paragraph."); // First boundary node.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("Middle paragraph.");
        builder.Writeln("End paragraph."); // Second boundary node.
        builder.Writeln("After paragraph.");
        source.Save("source.docx");

        // Step 2: Load the document from disk.
        Document loaded = new Document("source.docx");

        // Step 3: Locate the start and end paragraph nodes by their text content.
        Paragraph startPara = FindParagraphByText(loaded, "Start paragraph.");
        Paragraph endPara = FindParagraphByText(loaded, "End paragraph.");

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Step 4: Extract all nodes between the two paragraphs (inclusive) into a new document.
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure a clean document.

        // Build the minimal document structure: Section -> Body.
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        bool inRange = false;
        NodeCollection bodyChildren = loaded.FirstSection.Body.GetChildNodes(NodeType.Any, false);

        foreach (Node node in bodyChildren)
        {
            if (node == startPara)
                inRange = true;

            if (inRange)
            {
                // Import the node so it belongs to the destination document.
                Node importedNode = importer.ImportNode(node, true);

                // Append only block-level nodes directly to the body.
                if (importedNode.NodeType == NodeType.Paragraph || importedNode.NodeType == NodeType.Table)
                {
                    resultBody.AppendChild(importedNode);
                }
                else
                {
                    // For any inline nodes (unlikely at this level), wrap them in a paragraph.
                    Paragraph wrapper = new Paragraph(result);
                    wrapper.AppendChild(importedNode);
                    resultBody.AppendChild(wrapper);
                }
            }

            if (node == endPara)
                break;
        }

        // Step 5: Encrypt the extracted document with a password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "Secret123"
        };
        string outputPath = "extracted_encrypted.docx";
        result.Save(outputPath, saveOptions);

        // Step 6: Validate that the encrypted file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Encrypted output file was not created.");

        // Optional verification: load the encrypted file with the password and check content.
        Document encryptedDoc = new Document(outputPath, new LoadOptions("Secret123"));
        string extractedText = encryptedDoc.GetText();

        if (!extractedText.Contains("Start paragraph.") ||
            !extractedText.Contains("End paragraph.") ||
            !extractedText.Contains("Cell 1"))
        {
            throw new InvalidOperationException("Extracted content validation failed.");
        }

        Console.WriteLine("Extraction and encryption completed successfully.");
    }

    // Helper: finds the first paragraph whose text (trimmed) matches the supplied string.
    private static Paragraph FindParagraphByText(Document doc, string textToFind)
    {
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // GetText includes the paragraph break; Trim removes whitespace and line breaks.
            if (para.GetText().Trim() == textToFind)
                return para;
        }
        return null;
    }
}
