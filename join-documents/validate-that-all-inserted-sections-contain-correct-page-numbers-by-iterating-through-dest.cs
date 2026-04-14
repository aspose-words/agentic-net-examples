using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create first source document ----------
        Document srcDoc1 = new Document();
        DocumentBuilder srcBuilder1 = new DocumentBuilder(srcDoc1);
        srcBuilder1.Writeln("Content of Source Document 1");
        srcBuilder1.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        srcBuilder1.InsertField(FieldType.FieldPage, true);
        string srcPath1 = Path.Combine(outputDir, "Source1.docx");
        srcDoc1.Save(srcPath1, SaveFormat.Docx);

        // ---------- Create second source document ----------
        Document srcDoc2 = new Document();
        DocumentBuilder srcBuilder2 = new DocumentBuilder(srcDoc2);
        srcBuilder2.Writeln("Content of Source Document 2");
        srcBuilder2.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        srcBuilder2.InsertField(FieldType.FieldPage, true);
        string srcPath2 = Path.Combine(outputDir, "Source2.docx");
        srcDoc2.Save(srcPath2, SaveFormat.Docx);

        // ---------- Create destination document ----------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Content of Destination Document");
        dstBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        dstBuilder.InsertField(FieldType.FieldPage, true);

        // Append the source documents.
        dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Update fields and layout so page numbers are calculated.
        dstDoc.UpdateFields();
        dstDoc.UpdatePageLayout();

        // ---------- Validate page numbers in each section's footer ----------
        for (int i = 0; i < dstDoc.Sections.Count; i++)
        {
            Section section = dstDoc.Sections[i];
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer == null)
                throw new InvalidOperationException($"Section {i + 1} does not contain a primary footer.");

            // Expect exactly one PAGE field in the footer.
            Field? pageField = null;
            foreach (Field field in footer.Range.Fields)
            {
                if (field.Type == FieldType.FieldPage)
                {
                    pageField = field;
                    break;
                }
            }

            if (pageField == null)
                throw new InvalidOperationException($"Footer of section {i + 1} does not contain a PAGE field.");

            // The field result contains the page number as text.
            string resultText = pageField.Result.Trim();
            if (!int.TryParse(resultText, out int pageNumber))
                throw new InvalidOperationException($"Footer of section {i + 1} contains a non‑numeric page number \"{resultText}\".");

            int expectedPage = i + 1; // Sections start on consecutive pages.
            if (pageNumber != expectedPage)
                throw new InvalidOperationException($"Section {i + 1} has page number {pageNumber}, expected {expectedPage}.");
        }

        // Save the merged document.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        dstDoc.Save(mergedPath, SaveFormat.Docx);
    }
}
