using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveTocExample
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("InputDocument.doc");

        // Find all Table of Contents (TOC) fields and remove them.
        // FieldToc is the correct class name in Aspose.Words.
        foreach (FieldToc toc in doc.Range.Fields.OfType<FieldToc>().ToList())
        {
            toc.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save("OutputDocument.docx", SaveFormat.Docx);
    }
}
