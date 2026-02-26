using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC file that contains a table of contents.
            string inputFile = @"C:\Docs\SourceDocument.doc";

            // Path where the resulting DOC file without the TOC will be saved.
            string outputFile = @"C:\Docs\DocumentWithoutToc.doc";

            // Load the existing document.
            Document doc = new Document(inputFile);

            // Collect all TOC fields in the document.
            List<FieldToc> tocFields = new List<FieldToc>();
            foreach (Field field in doc.Range.Fields)
            {
                if (field is FieldToc toc)
                {
                    tocFields.Add(toc);
                }
            }

            // Remove each TOC field from the document.
            foreach (FieldToc toc in tocFields)
            {
                toc.Remove();
            }

            // Save the modified document back to DOC format.
            doc.Save(outputFile);
        }
    }
}
