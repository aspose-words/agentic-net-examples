using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string docPath = Path.Combine(outputDir, "OfficeMath.docx");
        string pdfPath = Path.Combine(outputDir, "OfficeMath.pdf");

        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before equations.
        builder.Writeln("Sample document with OfficeMath equations:");

        // Insert first EQ field (fraction 1/2).
        InsertFieldEQ(builder, @"\f(1,2)");

        // Insert second EQ field (square root of x).
        InsertFieldEQ(builder, @"\r(2,x)");

        // Convert all EQ fields to real OfficeMath objects.
        var eqFields = doc.Range.Fields.OfType<FieldEQ>().ToList();
        foreach (FieldEQ field in eqFields)
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

        // Adjust display type and justification for top‑level equations.
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Verify that the expected number of top‑level equations exists.
        int expectedEquations = 2;
        int actualEquations = officeMathNodes
            .Cast<OfficeMath>()
            .Count(om => om.MathObjectType == MathObjectType.OMathPara);

        if (actualEquations != expectedEquations)
            throw new InvalidOperationException($"Expected {expectedEquations} equations, but found {actualEquations}.");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);
    }

    // Helper method to insert an EQ field with the specified argument string.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return to the field start's parent node and start a new paragraph.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
        return field;
    }
}
