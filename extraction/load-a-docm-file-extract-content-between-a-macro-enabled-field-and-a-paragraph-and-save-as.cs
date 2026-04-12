using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class ExtractBetweenMacroFieldAndParagraph
{
    public static void Main()
    {
        // Create a temporary folder for the sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeExtractDemo");
        Directory.CreateDirectory(tempFolder);

        // Paths for the source DOCM and the result DOCX.
        string sourceDocPath = Path.Combine(tempFolder, "SourceDocument.docm");
        string resultDocPath = Path.Combine(tempFolder, "ExtractedContent.docx");

        // -------------------------------------------------
        // 1. Build a sample macro‑enabled document (DOCM).
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a MACROBUTTON field.
        FieldMacroButton macroField = (FieldMacroButton)builder.InsertField(FieldType.FieldMacroButton, true);
        macroField.MacroName = "SampleMacro";
        macroField.DisplayText = "Run SampleMacro";

        // Add some content after the macro field.
        builder.Writeln(); // move to a new paragraph.
        builder.Writeln("Paragraph 1: after macro field.");
        builder.Writeln("Paragraph 2: more content.");
        builder.Writeln("TargetParagraph: this is the paragraph that marks the end of extraction.");

        // Save as a macro‑enabled document.
        sourceDoc.Save(sourceDocPath, SaveFormat.Docm);

        // -------------------------------------------------
        // 2. Load the DOCM document.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);

        // -------------------------------------------------
        // 3. Locate the macro button field.
        // -------------------------------------------------
        FieldMacroButton startField = null;
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldMacroButton)
            {
                startField = (FieldMacroButton)field;
                break;
            }
        }

        if (startField == null)
            throw new InvalidOperationException("Macro button field not found in the document.");

        // -------------------------------------------------
        // 4. Locate the target paragraph (by its unique text).
        // -------------------------------------------------
        Paragraph endParagraph = null;
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("TargetParagraph"))
            {
                endParagraph = para;
                break;
            }
        }

        if (endParagraph == null)
            throw new InvalidOperationException("Target paragraph not found in the document.");

        // -------------------------------------------------
        // 5. Extract nodes that lie between the field and the paragraph.
        // -------------------------------------------------
        // The field resides inside a paragraph. The first node to extract is the paragraph
        // that follows the field's parent paragraph.
        Paragraph startParagraph = (Paragraph)startField.End.ParentNode;
        Node currentNode = startParagraph.NextSibling; // first candidate node

        // Prepare the destination document.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // clear the default empty section/paragraph

        // Create a new section with a body.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Use NodeImporter for proper style and list handling.
        NodeImporter importer = new NodeImporter(loadedDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        while (currentNode != null && currentNode != endParagraph)
        {
            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(currentNode, true);

            // Append only nodes that are valid children of Body.
            if (importedNode.NodeType == NodeType.Paragraph ||
                importedNode.NodeType == NodeType.Table ||
                importedNode.NodeType == NodeType.Shape)
            {
                destBody.AppendChild(importedNode);
            }
            else if (importedNode.NodeType == NodeType.Run)
            {
                // Wrap isolated runs in a new paragraph.
                Paragraph wrapper = new Paragraph(destDoc);
                wrapper.AppendChild(importedNode);
                destBody.AppendChild(wrapper);
            }
            else
            {
                // For any other node types, attempt to append directly.
                destBody.AppendChild(importedNode);
            }

            currentNode = currentNode.NextSibling;
        }

        // -------------------------------------------------
        // 6. Validate that something was extracted.
        // -------------------------------------------------
        if (destBody.GetChildNodes(NodeType.Any, false).Count == 0)
            throw new InvalidOperationException("No content was extracted between the specified field and paragraph.");

        // -------------------------------------------------
        // 7. Save the extracted content as DOCX.
        // -------------------------------------------------
        destDoc.Save(resultDocPath, SaveFormat.Docx);

        // Inform the user where files are located (no interactive input required).
        Console.WriteLine("Source DOCM created at: " + sourceDocPath);
        Console.WriteLine("Extracted DOCX saved at: " + resultDocPath);
    }
}
