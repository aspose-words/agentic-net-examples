using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example data representing sections. In a real scenario this could come from a database,
        // a file, or any other source.
        List<string> sections = new List<string>
        {
            "Section 1 title\nSection 1 content.",
            "Section 2 title\nSection 2 content.",
            "Section 3 title\nSection 3 content."
        };

        // Outer foreach loop over the sections.
        foreach (string sectionText in sections)
        {
            // Write the section text.
            builder.Write(sectionText);

            // Insert a page break before the next section if the special tag <<pageBreak>> is present.
            // Here we simulate the presence of the tag by always inserting a break after each section
            // except the last one.
            if (sectionText != sections.Last())
            {
                // Insert an explicit page break.
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document to disk.
        doc.Save("Output.docx", SaveFormat.Docx);
    }
}
