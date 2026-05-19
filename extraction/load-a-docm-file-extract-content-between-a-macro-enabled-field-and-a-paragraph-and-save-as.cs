using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCM file that contains a macro button field and several paragraphs.
        string sourcePath = "sample.docm";
        CreateSampleDocm(sourcePath);

        // Load the DOCM document.
        Document sourceDoc = new Document(sourcePath);

        // Find the first MACROBUTTON field.
        Field macroField = null;
        foreach (Field field in sourceDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldMacroButton)
            {
                macroField = field;
                break;
            }
        }

        if (macroField == null)
            throw new InvalidOperationException("Macro button field not found.");

        // Find the paragraph that contains the text "Target paragraph".
        Paragraph targetParagraph = null;
        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("Target paragraph"))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Determine the start paragraph (the paragraph that holds the macro field).
        Paragraph startParagraph = macroField.Start.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Unable to locate the start paragraph.");

        // Build a new document that will receive the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Ensure a clean document.

        // Create a new section and body for the extracted document.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body extractedBody = new Body(extractedDoc);
        section.AppendChild(extractedBody);

        // Prepare an importer to copy nodes from the source document to the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk from the start paragraph to the target paragraph (inclusive) and import each node.
        Node currentNode = startParagraph;
        while (true)
        {
            // Import the current node (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(currentNode, true);
            extractedBody.AppendChild(importedNode);

            if (currentNode == targetParagraph)
                break;

            currentNode = currentNode.NextSibling;
            if (currentNode == null)
                throw new InvalidOperationException("Reached end of document before finding the target paragraph.");
        }

        // Save the extracted content as a DOCX file.
        string outputPath = "extracted.docx";
        extractedDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");
    }

    private static void CreateSampleDocm(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a macro button field.
        builder.InsertField("MACROBUTTON NoMacro \"ClickMe\"");

        // Add some content after the field.
        builder.Writeln();
        builder.Writeln("Content start");
        builder.Writeln("Content middle");
        builder.Writeln("Target paragraph");

        // Save as a macro‑enabled document.
        doc.Save(filePath, SaveFormat.Docm);
    }
}
