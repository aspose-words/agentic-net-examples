using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create a sample DOCM file containing a macro button field and
        // a target paragraph that marks the end of the extraction range.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph.");

        // Insert a macro button field (MACROBUTTON). Use the string overload
        // to avoid overload ambiguity.
        builder.InsertField("MACROBUTTON NoMacro ClickMe");

        // Ensure the field is followed by a new paragraph.
        builder.Writeln();

        // Content that should be extracted.
        builder.Writeln("Content line 1.");
        builder.Writeln("Content line 2.");

        // Paragraph that marks the end boundary of the extraction.
        builder.Writeln("Target paragraph.");

        // Save the sample document as DOCM.
        sourceDoc.Save("sample.docm", SaveFormat.Docm);

        // ------------------------------------------------------------
        // Load the DOCM file and locate the macro button field and the
        // target paragraph.
        // ------------------------------------------------------------
        Document loadedDoc = new Document("sample.docm");

        Field macroField = null;
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldMacroButton)
            {
                macroField = field;
                break;
            }
        }

        if (macroField == null)
            throw new InvalidOperationException("Macro button field not found.");

        Paragraph targetParagraph = null;
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Trim() == "Target paragraph.")
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // ------------------------------------------------------------
        // Build a new document that will contain the extracted nodes.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // ------------------------------------------------------------
        // Clone nodes that lie between the macro field end and the target paragraph.
        // ------------------------------------------------------------
        Node currentNode = macroField.End.NextSibling;
        while (currentNode != null && currentNode != targetParagraph)
        {
            // Clone the node deeply and append it to the result body.
            Node clonedNode = currentNode.Clone(true);
            resultBody.AppendChild(clonedNode);
            currentNode = currentNode.NextSibling;
        }

        // ------------------------------------------------------------
        // Save the extracted content as DOCX and validate the output.
        // ------------------------------------------------------------
        string outputPath = "extracted.docx";
        resultDoc.Save(outputPath, SaveFormat.Docx);

        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");
    }
}
