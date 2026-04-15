using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string sourceDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Exported.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a top‑level OfficeMath equation.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some surrounding text.
        builder.Writeln("The following equation is displayed on its own line:");

        // Insert an EQ field that will be converted to a real OfficeMath node.
        // Use a simple, deterministic equation: a fraction 1/2.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that its internal state reflects the new code.
        eqField.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Ensure the equation is a top‑level paragraph and set its display properties.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display; // Display on its own line.
            officeMath.Justification = OfficeMathJustification.Left; // Left‑aligned.
        }

        // Save the source DOCX.
        doc.Save(sourceDocPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 2. Export the document to PDF, enabling additional text positioning.
        // ---------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AdditionalTextPositioning = true // Improves positioning accuracy.
        };
        doc.Save(pdfPath, pdfOptions);

        // ---------------------------------------------------------------
        // 3. Validation.
        // ---------------------------------------------------------------
        // Verify that both files were created.
        if (!File.Exists(sourceDocPath))
            throw new FileNotFoundException("Source DOCX was not created.", sourceDocPath);
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Exported PDF was not created.", pdfPath);

        // Load the source document again to inspect OfficeMath nodes.
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Count only top‑level OfficeMath paragraphs (OMathPara).
        int topLevelCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelCount++;
        }

        // Expect exactly one top‑level equation.
        if (topLevelCount != 1)
            throw new InvalidOperationException($"Expected 1 top‑level OfficeMath node, but found {topLevelCount}.");

        // If we reach this point, the PDF was generated and the source equation
        // is correctly positioned as a display equation.
        Console.WriteLine("Validation succeeded: source DOCX contains the expected OfficeMath equation and PDF was created.");
    }
}
