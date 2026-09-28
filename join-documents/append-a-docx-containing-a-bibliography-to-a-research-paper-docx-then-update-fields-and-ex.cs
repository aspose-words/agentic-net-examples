using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample documents and the final PDF.
        string researchPath = "ResearchPaper.docx";
        string bibliographyPath = "Bibliography.docx";
        string outputPdfPath = "MergedDocument.pdf";

        // -----------------------------------------------------------------
        // Create a sample research paper DOCX.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the introduction of the research paper.");
        // Insert a simple PAGE field to demonstrate field updating later.
        researchBuilder.InsertField("PAGE", "1");
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography DOCX.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        bibBuilder.Writeln("1. Author A. Title A.");
        bibBuilder.Writeln("2. Author B. Title B.");
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the research paper and append the bibliography.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(researchPath);
        Document bibToAppend = new Document(bibliographyPath);
        mainDoc.AppendDocument(bibToAppend, ImportFormatMode.KeepSourceFormatting);

        // Update all fields (e.g., PAGE fields) after the merge.
        mainDoc.UpdateFields();

        // Save the merged document as PDF.
        mainDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
        {
            throw new InvalidOperationException("The merged PDF was not created correctly.");
        }

        // Optional: clean up sample DOCX files (comment out if inspection is needed).
        // File.Delete(researchPath);
        // File.Delete(bibliographyPath);
    }
}
