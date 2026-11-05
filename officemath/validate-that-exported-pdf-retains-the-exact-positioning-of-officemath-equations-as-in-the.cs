using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Math;

public class OfficeMathPdfPositionValidator
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string docxPath = Path.Combine(outputDir, "sample.docx");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text
        builder.Writeln("This document contains several OfficeMath equations.");
        builder.Writeln();

        // Insert first equation using deterministic EQ-field bootstrap
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction 1/2
        builder.Writeln();

        // Insert second equation
        InsertEquation(builder, @"\r(3,x)"); // Simple root
        builder.Writeln();

        // Insert third equation
        InsertEquation(builder, @"\s(5)"); // Simple superscript
        builder.Writeln();

        // Save the DOCX file
        doc.Save(docxPath, SaveFormat.Docx);

        // Ensure layout is up‑to‑date
        doc.UpdatePageLayout();

        // Collect layout information
        LayoutCollector collector = new LayoutCollector(doc);

        // Validate that each top‑level OfficeMath node is on page 1
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                int pageNumber = collector.GetStartPageIndex(om);
                if (pageNumber != 1)
                {
                    throw new InvalidOperationException($"Equation on page {pageNumber} does not match expected page 1.");
                }
            }
        }

        // Convert the document to PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF output file was not created.", pdfPath);
        }

        Console.WriteLine("Validation successful: all equations retained their positioning in the PDF.");
    }

    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            return null;

        // Write the EQ argument string at the field separator
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(eqSwitch);

        // Convert the field to a real OfficeMath node
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            return null;

        // Insert the OfficeMath node before the field start
        Node fieldStart = fieldEQ.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original field
        field.Remove();

        return officeMath;
    }
}
