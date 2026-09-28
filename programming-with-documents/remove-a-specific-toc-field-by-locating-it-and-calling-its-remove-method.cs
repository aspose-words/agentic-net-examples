using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add some headings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents field.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.Writeln(); // Add a paragraph break after TOC.

            // Add headings that will appear in the TOC.
            for (int i = 1; i <= 3; i++)
            {
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Heading {i}");
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln($"Content for heading {i}.");
            }

            // Save the original document.
            const string originalPath = "Original.docx";
            doc.Save(originalPath);

            // Load the document to locate and remove the TOC field.
            Document loadedDoc = new Document(originalPath);

            // Find the first TOC field and remove it.
            foreach (Field field in loadedDoc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    field.Remove();
                    break; // Remove only the specific TOC field.
                }
            }

            // Save the modified document.
            const string modifiedPath = "Modified.docx";
            loadedDoc.Save(modifiedPath);
        }
    }
}
