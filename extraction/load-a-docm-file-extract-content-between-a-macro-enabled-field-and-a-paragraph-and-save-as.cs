using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class ExtractionExample
{
    public static void Main()
    {
        // Create a sample DOCM file with a MacroButton field, some content, and an end marker paragraph.
        string sourcePath = "source.docm";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a MacroButton field.
        builder.InsertField(@"MACROBUTTON ""MyMacro"" ""ClickMe""", "ClickMe");
        builder.Writeln(); // End the paragraph containing the field.

        // Add content that should be extracted.
        builder.Writeln("First extracted paragraph.");
        builder.Writeln("Second extracted paragraph.");

        // End marker paragraph.
        builder.Writeln("EndMarkerParagraph");

        // Save as DOCM.
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // Load the DOCM file.
        Document loadedDoc = new Document(sourcePath);

        // Locate the MacroButton field start.
        FieldStart macroFieldStart = null;
        foreach (Node node in loadedDoc.GetChildNodes(NodeType.FieldStart, true))
        {
            FieldStart fs = node as FieldStart;
            if (fs != null && fs.FieldType == FieldType.FieldMacroButton)
            {
                macroFieldStart = fs;
                break;
            }
        }

        if (macroFieldStart == null)
            throw new InvalidOperationException("MacroButton field not found.");

        // Locate the paragraph that contains the end marker.
        Paragraph endParagraph = null;
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("EndMarkerParagraph"))
            {
                endParagraph = para;
                break;
            }
        }

        if (endParagraph == null)
            throw new InvalidOperationException("End marker paragraph not found.");

        // Determine the paragraph that contains the macro field.
        Paragraph startParagraph = macroFieldStart.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph containing the macro field not found.");

        // Collect paragraphs between startParagraph (exclusive) and endParagraph (exclusive).
        List<Paragraph> extractedParagraphs = new List<Paragraph>();
        Node current = startParagraph.NextSibling;
        while (current != null && current != endParagraph)
        {
            if (current.NodeType == NodeType.Paragraph)
                extractedParagraphs.Add((Paragraph)current);
            current = current.NextSibling;
        }

        if (extractedParagraphs.Count == 0)
            throw new InvalidOperationException("No content found between the macro field and the end paragraph.");

        // Create a new document and copy the extracted paragraphs.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);
        Body body = new Body(resultDoc);
        section.AppendChild(body);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Paragraph para in extractedParagraphs)
        {
            Node importedNode = importer.ImportNode(para, true);
            body.AppendChild(importedNode);
        }

        // Save the extracted content as DOCX.
        string outputPath = "extracted.docx";
        resultDoc.Save(outputPath, SaveFormat.Docx);

        // Validate output.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");
    }
}
