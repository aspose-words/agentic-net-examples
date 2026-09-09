using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCM file with a macro button field, some content, and a terminating paragraph.
        string sourcePath = "sample.docm";
        CreateSampleDocm(sourcePath);

        // Load the DOCM file.
        Document sourceDoc = new Document(sourcePath);

        // Locate the macro button field.
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

        // Locate the target paragraph that marks the end of the extraction range.
        Paragraph endParagraph = null;
        foreach (Paragraph para in sourceDoc.FirstSection.Body.Paragraphs)
        {
            if (para.GetText().Contains("End Paragraph"))
            {
                endParagraph = para;
                break;
            }
        }

        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // Determine the paragraph that contains the macro field.
        Paragraph startParagraph = macroField.Start.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // Build a new document that will hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);

        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Get the collection of paragraphs in the source body.
        ParagraphCollection sourceParas = sourceDoc.FirstSection.Body.Paragraphs;

        // Find the indices of the start and end paragraphs.
        int startIndex = sourceParas.IndexOf(startParagraph);
        int endIndex = sourceParas.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph range for extraction.");

        // Use a NodeImporter to import nodes from the source document into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Import and copy each paragraph within the range to the new document.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Node importedNode = importer.ImportNode(sourceParas[i], true);
            body.AppendChild(importedNode);
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
        builder.InsertField("MACROBUTTON NoMacro \"Click Me\"");

        // Add some content after the field.
        builder.Writeln("Content line 1");
        builder.Writeln("Content line 2");

        // Insert the terminating paragraph.
        builder.Writeln("End Paragraph");

        // Save as a macro-enabled document.
        doc.Save(filePath, SaveFormat.Docm);
    }
}
