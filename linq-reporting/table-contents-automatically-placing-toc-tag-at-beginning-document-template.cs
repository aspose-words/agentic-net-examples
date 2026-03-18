using System;
using Aspose.Words;
using Aspose.Words.Saving;

class GenerateToc
{
    static void Main()
    {
        // Path to the existing document template.
        string templatePath = @"C:\Docs\Template.docx";

        // Path where the resulting document will be saved.
        string outputPath = @"C:\Docs\DocumentWithToc.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a TOC field. The switches configure the TOC to:
        //   - include heading levels 1 through 3 (\o "1-3")
        //   - make entries clickable hyperlinks (\h)
        //   - hide page numbers in web layout (\z)
        //   - use outline levels from heading styles (\u)
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Optionally insert a page break after the TOC so that content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Update all fields in the document so the TOC reflects the current headings.
        doc.UpdateFields();

        // Save the document. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
