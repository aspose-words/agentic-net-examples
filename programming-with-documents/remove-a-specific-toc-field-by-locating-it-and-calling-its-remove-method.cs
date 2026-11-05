using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Chapter 2");

        // Insert the Table of Contents at the beginning of the document.
        builder.MoveToDocumentStart();
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Save the original document.
        const string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document (could also continue using the same instance).
        Document loadedDoc = new Document(originalPath);

        // Locate the first TOC field in the document.
        Field tocField = null;
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                tocField = field;
                break;
            }
        }

        // Remove the TOC field if it was found.
        tocField?.Remove();

        // Save the modified document.
        const string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);
    }
}
