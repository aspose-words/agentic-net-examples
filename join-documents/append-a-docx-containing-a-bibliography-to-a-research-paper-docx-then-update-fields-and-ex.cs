using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // File names for the sample documents and the merged results.
        const string researchPath = "ResearchPaper.docx";
        const string bibliographyPath = "Bibliography.docx";
        const string mergedDocPath = "Merged.docx";
        const string mergedPdfPath = "Merged.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample research paper document.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the content of the research paper.");
        researchBuilder.InsertBreak(BreakType.PageBreak);
        // Insert a PAGE field that will be updated after the merge.
        researchBuilder.InsertField(FieldType.FieldPage, true);
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a sample bibliography document.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        // Insert a BIBLIOGRAPHY field.
        bibBuilder.InsertField(FieldType.FieldBibliography, true);
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the two documents.
        // -----------------------------------------------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // -----------------------------------------------------------------
        // 4. Append the bibliography to the research paper.
        // -----------------------------------------------------------------
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Update all fields in the combined document.
        // -----------------------------------------------------------------
        research.UpdateFields();

        // -----------------------------------------------------------------
        // 6. Save the merged document as DOCX and PDF.
        // -----------------------------------------------------------------
        research.Save(mergedDocPath, SaveFormat.Docx);
        research.Save(mergedPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 7. Validation: ensure the PDF exists and bibliography text is present.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPdfPath))
            throw new InvalidOperationException("The merged PDF was not created.");

        string mergedText = research.GetText();
        if (!mergedText.Contains("Bibliography"))
            throw new InvalidOperationException("The bibliography content was not appended.");
    }
}
