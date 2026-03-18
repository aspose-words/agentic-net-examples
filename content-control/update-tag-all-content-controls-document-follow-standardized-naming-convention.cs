using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Retrieve the collection of all structured document tags (content controls) in the document.
        StructuredDocumentTagCollection sdtCollection = doc.Range.StructuredDocumentTags;

        // Update each tag to follow a standardized naming convention.
        // Example convention: "CC_{TitleOrIndex}_{ZeroBasedIndex}"
        for (int i = 0; i < sdtCollection.Count; i++)
        {
            IStructuredDocumentTag sdt = sdtCollection[i];

            // Use the existing Title (if any) as part of the new tag, otherwise fall back to a generic name.
            string titlePart = !string.IsNullOrEmpty(sdt.Title)
                ? sdt.Title.Replace(' ', '_')
                : $"Tag{i}";

            // Assign the new standardized tag.
            sdt.Tag = $"CC_{titlePart}_{i}";
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
