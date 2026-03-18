using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace TocPageNumberRebuilder
{
    class Program
    {
        static void Main()
        {
            // Load an existing document that contains a Table of Contents.
            Document doc = new Document("Input.docx");

            // First, update all fields that are not dependent on page layout
            // (e.g., REF, DATE, etc.). This ensures the TOC field itself is up‑to‑date.
            doc.UpdateFields();

            // Rebuild the page layout so that page‑related fields (PAGE, NUMPAGES, etc.)
            // have correct values. This step is required before updating TOC page numbers.
            doc.UpdatePageLayout();

            // Iterate through all fields in the document and find TOC fields.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // Cast to FieldToc to access the UpdatePageNumbers method.
                    FieldToc toc = (FieldToc)field;
                    // Update the page numbers displayed in this TOC.
                    toc.UpdatePageNumbers();
                }
            }

            // Save the updated document.
            doc.Save("Output.docx");
        }
    }
}
