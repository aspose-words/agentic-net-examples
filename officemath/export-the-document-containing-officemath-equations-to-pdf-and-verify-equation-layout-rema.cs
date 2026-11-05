using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Path.GetTempPath(), "OfficeMathExample");
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "EquationDoc.docx");
        string pdfPath = Path.Combine(outputDir, "EquationDoc.pdf");

        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph introducing the equation
        builder.Writeln("Document with OfficeMath equation:");

        // Insert a field equation (EQ field) and write a simple fraction equation
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)"); // Represents the fraction 1/2

        // Update the field so that the EQ switch is processed
        field.Update();

        // Convert the EQ field to a real OfficeMath node
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath officeMath = fieldEq?.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field
        Node fieldStart = field.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        field.Remove();

        // Set display type for the top‑level equation (only for OMathPara nodes)
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
            officeMath.DisplayType = OfficeMathDisplayType.Display;

        // Save the document as DOCX
        doc.Save(docxPath, SaveFormat.Docx);

        // Export the document to PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure PDF was created
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF file was not created.", pdfPath);

        // Validation: ensure the OfficeMath node exists and has the expected display type
        NodeCollection omNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (omNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes found in the document.");

        OfficeMath firstOm = (OfficeMath)omNodes[0];
        if (firstOm.MathObjectType == MathObjectType.OMathPara &&
            firstOm.DisplayType != OfficeMathDisplayType.Display)
        {
            throw new InvalidOperationException("OfficeMath display type does not match the expected value.");
        }

        // Indicate success (non‑interactive)
        Console.WriteLine("OfficeMath equation exported to PDF successfully.");
    }
}
