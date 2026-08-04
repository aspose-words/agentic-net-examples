using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCM file with a field and several paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph 1 – contains a macro‑enabled field (for demo we use a simple field).
        builder.InsertField("MERGEFIELD SampleField \\* MERGEFORMAT");
        builder.Writeln(); // End of first paragraph.

        // Paragraph 2 – content that will be extracted.
        builder.Writeln("This is the content that should be extracted.");

        // Paragraph 3 – the ending paragraph (boundary).
        builder.Writeln("End of extraction range.");

        // Save as a macro‑enabled document (DOCM).
        const string sourcePath = "sample.docm";
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the DOCM file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the field (start marker) and the ending paragraph.
        // -----------------------------------------------------------------
        if (loadedDoc.Range.Fields.Count == 0)
            throw new InvalidOperationException("No fields found in the document.");

        // Assume the first field is the start marker.
        Field startField = loadedDoc.Range.Fields[0];
        Paragraph startParagraph = startField.Start.ParentNode as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Start field is not inside a paragraph.");

        // Find the ending paragraph by its exact text.
        Paragraph endParagraph = null;
        foreach (Paragraph para in loadedDoc.FirstSection.Body.Paragraphs)
        {
            if (para.GetText().Trim() == "End of extraction range.")
            {
                endParagraph = para;
                break;
            }
        }
        if (endParagraph == null)
            throw new InvalidOperationException("Ending paragraph not found.");

        // -----------------------------------------------------------------
        // 4. Extract the content that lies between the start field's paragraph
        //    and the ending paragraph (exclusive of the boundaries).
        // -----------------------------------------------------------------
        Body sourceBody = loadedDoc.FirstSection.Body;
        int startIndex = sourceBody.Paragraphs.IndexOf(startParagraph);
        int endIndex = sourceBody.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex + 1)
            throw new InvalidOperationException("Invalid extraction range.");

        // Create a new document that will hold the extracted content.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove the default section/paragraph.

        // Build the minimal required structure: Section -> Body.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each paragraph that falls inside the range.
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            Paragraph paraToImport = sourceBody.Paragraphs[i];
            Node importedNode = importer.ImportNode(paraToImport, true);
            resultBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted content as a DOCX file.
        // -----------------------------------------------------------------
        const string outputPath = "extracted.docx";
        resultDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");
    }
}
