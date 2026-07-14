using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathPdfExport
{
    public static void Main()
    {
        // Prepare output folder
        string outputFolder = Path.Combine(Path.GetTempPath(), "OfficeMathExample");
        Directory.CreateDirectory(outputFolder);

        string docxPath = Path.Combine(outputFolder, "EquationDocument.docx");
        string pdfPath = Path.Combine(outputFolder, "EquationDocument.pdf");

        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add introductory text
        builder.Writeln("Below is an OfficeMath equation:");

        // Insert a field that will be converted to OfficeMath
        builder.InsertField(FieldType.FieldEquation, true);
        // Retrieve the inserted field (the last field in the document)
        FieldEQ eqField = doc.Range.Fields[doc.Range.Fields.Count - 1] as FieldEQ;
        if (eqField == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a simple EQ argument into the field separator
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // Simple fraction 1/2

        // Convert the field to a real OfficeMath node
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field
        Node fieldStart = eqField.Start;
        Paragraph parentParagraph = fieldStart.ParentNode as Paragraph;
        if (parentParagraph == null)
            throw new InvalidOperationException("Field start does not have a paragraph parent.");

        parentParagraph.InsertBefore(officeMath, fieldStart);
        eqField.Remove();

        // Validate that the document now contains exactly one top‑level OfficeMath paragraph
        var topLevelMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                              .Cast<OfficeMath>()
                              .Where(om => om.MathObjectType == MathObjectType.OMathPara)
                              .ToList();

        if (topLevelMath.Count != 1)
            throw new InvalidOperationException($"Expected 1 top‑level OfficeMath node, found {topLevelMath.Count}.");

        // Save the document as DOCX
        doc.Save(docxPath, SaveFormat.Docx);

        // Export the document to PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that both output files exist
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("DOCX output file was not created.", docxPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF output file was not created.", pdfPath);

        // Example completed successfully
        Console.WriteLine("Document and PDF have been created successfully:");
        Console.WriteLine($"DOCX: {docxPath}");
        Console.WriteLine($"PDF: {pdfPath}");
    }
}
