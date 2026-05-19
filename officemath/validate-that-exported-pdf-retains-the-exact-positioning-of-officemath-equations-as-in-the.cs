using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathPdfPositionValidation
{
    public static void Main()
    {
        // Prepare output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "SampleOfficeMath.docx");
        string pdfPath = Path.Combine(outputDir, "SampleOfficeMath.pdf");

        // 1. Create a new document and insert two simple OfficeMath equations using the EQ‑field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation: a simple fraction 1/2.
        InsertEquation(builder, @"\f(1,2)");
        // Second equation: a cubic root of x.
        InsertEquation(builder, @"\r(3,x)");

        // 2. Convert all inserted EQ fields to real OfficeMath nodes and remove the original fields.
        foreach (FieldEQ field in doc.Range.Fields.OfType<FieldEQ>().ToList())
        {
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }
        }

        // 3. Validate that the document contains exactly two top‑level OfficeMath paragraphs.
        int topLevelMathCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                                   .Cast<OfficeMath>()
                                   .Count(m => m.MathObjectType == MathObjectType.OMathPara);
        if (topLevelMathCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level OfficeMath nodes, but found {topLevelMathCount}.");

        // 4. Save the DOCX file.
        doc.Save(docPath, SaveFormat.Docx);
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the DOCX file.", docPath);

        // 5. Export the document to PDF with additional text positioning enabled to preserve exact layout.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AdditionalTextPositioning = true
        };
        doc.Save(pdfPath, pdfOptions);

        // 6. Validate that the PDF file was created and has non‑zero size.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF file.", pdfPath);
        if (new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("The generated PDF file is empty.");

        // All validations passed.
        Console.WriteLine("PDF export validation succeeded.");
    }

    // Helper method that inserts an EQ field with the given argument string,
    // then moves the builder to a new paragraph ready for the next insertion.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the field start's parent (the paragraph) and start a new paragraph.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
    }
}
