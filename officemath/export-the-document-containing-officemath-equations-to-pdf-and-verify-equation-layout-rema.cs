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
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "Equations.docx");
        string pdfPath = Path.Combine(outputDir, "Equations.pdf");

        // 1. Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Introductory paragraph.
        builder.Writeln("Below are sample equations created via EQ fields:");

        // 3. Insert two equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // fraction 1/2
        InsertEquation(builder, @"\r(3,x)"); // cube root of x

        // 4. Verify that two top‑level OfficeMath nodes (Math paragraphs) exist.
        int topLevelOfficeMathCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                                          .Cast<OfficeMath>()
                                          .Count(om => om.MathObjectType == MathObjectType.OMathPara);
        if (topLevelOfficeMathCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level OfficeMath nodes, but found {topLevelOfficeMathCount}.");

        // 5. Save the document as DOCX (optional).
        doc.Save(docPath, SaveFormat.Docx);

        // 6. Export the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // 7. Verify that the PDF file was created and is not empty.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF file was not created correctly.");

        // 8. Load the PDF back and ensure it no longer contains OfficeMath nodes
        //    (PDF conversion renders equations as images).
        Document pdfDoc = new Document(pdfPath);
        int pdfOfficeMathCount = pdfDoc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (pdfOfficeMathCount != 0)
            throw new InvalidOperationException($"PDF should not contain OfficeMath nodes, but found {pdfOfficeMathCount}.");

        // All validations passed.
        Console.WriteLine("Document and PDF generated successfully.");
    }

    // Inserts an EQ field, writes the equation code, updates the field,
    // converts it to a real OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqCode)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqCode);

        // Update the field so that its result (and OfficeMath conversion) is calculated.
        field.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Insert the OfficeMath node before the field start and remove the field.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
