using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled document (DOCM) from disk.
        // This uses the Document(string) constructor – a permitted lifecycle rule.
        Document sourceDoc = new Document("InputDocument.docm");

        // Find the first MACROBUTTON field in the document.
        // The field type for a macro button is FieldMacroButton.
        FieldMacroButton macroField = sourceDoc.Range.Fields
            .FirstOrDefault(f => f.Type == FieldType.FieldMacroButton) as FieldMacroButton;

        if (macroField == null)
        {
            Console.WriteLine("No MACROBUTTON field found in the source document.");
            return;
        }

        // The field end node marks the end of the macro field.
        Node fieldEnd = macroField.End;

        // Determine the paragraph that follows the macro field.
        // The field end is usually inside a paragraph; we get that paragraph and then its next sibling.
        Paragraph fieldParagraph = fieldEnd.GetAncestor(NodeType.Paragraph) as Paragraph;
        Paragraph nextParagraph = fieldParagraph?.NextSibling as Paragraph;

        if (nextParagraph == null)
        {
            Console.WriteLine("No paragraph found after the MACROBUTTON field.");
            return;
        }

        // Collect all nodes that lie between the field end and the next paragraph (exclusive).
        // This will include runs, tables, etc. that are directly after the field.
        var nodesToExtract = new System.Collections.Generic.List<Node>();
        Node curNode = fieldEnd.NextSibling;
        while (curNode != null && curNode != nextParagraph)
        {
            nodesToExtract.Add(curNode);
            curNode = curNode.NextSibling;
        }

        // Create a new blank document to hold the extracted content.
        // This uses the parameterless Document() constructor – a permitted lifecycle rule.
        Document extractedDoc = new Document();

        // Ensure the new document has at least one section and a body.
        extractedDoc.EnsureMinimum();

        // Import the collected nodes into the new document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            extractedDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Save the extracted content as a macro‑free DOCX file.
        // This uses the Document.Save(string) method – a permitted lifecycle rule.
        extractedDoc.Save("ExtractedContent.docx");

        Console.WriteLine("Extraction complete. Saved as ExtractedContent.docx");
    }
}
