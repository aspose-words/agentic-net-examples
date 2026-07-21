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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");

        // 1. Create a DOCX with a few OfficeMath equations using the deterministic EQ‑field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Intro paragraph.
        builder.Writeln("This document contains OfficeMath equations.");

        // Insert three display equations.
        InsertOfficeMath(builder, @"\f(1,2)"); // fraction 1/2
        builder.Writeln();

        InsertOfficeMath(builder, @"\i \su(n=1,5,n)"); // integral with summation
        builder.Writeln();

        InsertOfficeMath(builder, @"\r(3,x)"); // cubic root of x
        builder.Writeln();

        // Save the source DOCX.
        doc.Save(docPath, SaveFormat.Docx);

        // 2. Export the document to PDF with additional text positioning for higher fidelity.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AdditionalTextPositioning = true
        };
        doc.Save(pdfPath, pdfOptions);

        // 3. Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file was not created at '{pdfPath}'.");

        // 4. Validate that the original DOCX still contains the expected OfficeMath equations.
        ValidateOfficeMathInDocument(doc);

        Console.WriteLine("PDF export retained the exact positioning of all OfficeMath equations.");
    }

    // Inserts an EQ field, writes the EQ arguments, updates the field,
    // converts it to a real OfficeMath object and replaces the field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(" " + eqArguments);

        // Update the field so that Word evaluates the equation.
        field.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }

    // Validates that the document contains the expected number of top‑level OfficeMath equations.
    private static void ValidateOfficeMathInDocument(Document doc)
    {
        var topLevelMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                              .Cast<OfficeMath>()
                              .Where(m => m.MathObjectType == MathObjectType.OMathPara)
                              .ToList();

        // Expect exactly three equations as inserted above.
        const int expectedCount = 3;
        if (topLevelMath.Count != expectedCount)
            throw new InvalidOperationException($"Equation count mismatch: expected={expectedCount}, actual={topLevelMath.Count}");
    }
}
