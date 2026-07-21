using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample macro‑enabled document (DOCM) and save it.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a MACROBUTTON field (acts as a macro‑enabled field).
        builder.InsertField("MACROBUTTON NoMacro \"ClickMe\"");

        // Add some content that will be extracted.
        builder.Writeln("First line after macro field.");
        builder.Writeln("Second line after macro field.");

        // Paragraph that marks the end of the extraction range.
        builder.Writeln("Target paragraph.");

        const string sourcePath = "sample.docm";
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the DOCM file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the macro button field.
        Field macroField = null;
        foreach (Field f in loadedDoc.Range.Fields)
        {
            if (f.GetFieldCode().Contains("MACROBUTTON"))
            {
                macroField = f;
                break;
            }
        }

        if (macroField == null)
            throw new InvalidOperationException("Macro button field not found.");

        // Find the paragraph that contains the field.
        Paragraph fieldParagraph = macroField.Start.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (fieldParagraph == null)
            throw new InvalidOperationException("Field paragraph not found.");

        // Find the target paragraph (the one that contains the text "Target paragraph.").
        Paragraph targetParagraph = null;
        foreach (Paragraph p in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (p.GetText().Contains("Target paragraph."))
            {
                targetParagraph = p;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // -----------------------------------------------------------------
        // 3. Build a new document that will contain the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 4. Copy paragraphs from the source document to the result document.
        //    Use ImportNode to transfer nodes between documents safely.
        // -----------------------------------------------------------------
        Paragraph current = fieldParagraph.NextSibling as Paragraph;
        while (current != null)
        {
            // Import the paragraph (deep clone) into the destination document.
            Node imported = resultDoc.ImportNode(current, true);
            resultBody.AppendChild(imported);

            // Stop after adding the target paragraph.
            if (current == targetParagraph)
                break;

            current = current.NextSibling as Paragraph;
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted content as DOCX and verify the file.
        // -----------------------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath, SaveFormat.Docx);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");
    }
}
