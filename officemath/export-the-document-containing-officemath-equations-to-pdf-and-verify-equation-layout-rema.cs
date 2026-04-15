using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ExportOfficeMathToPdf
{
    public static void Main()
    {
        // Output file names
        const string docPath = "OfficeMathSample.docx";
        const string pdfPath = "OfficeMathSample.pdf";

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph
        builder.Writeln("Below is a sample equation created via EQ field bootstrap:");

        // Insert a few safe equations using the deterministic EQ‑field bootstrap workflow
        InsertEquation(builder, @"\f(1,2)");   // simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");   // cube root of x
        InsertEquation(builder, @"\i(, ,\f(x,2))"); // integral with a fraction

        // Set display mode and justification for top‑level equations only
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the intermediate DOCX (optional, useful for inspection)
        doc.Save(docPath, SaveFormat.Docx);

        // Export the document to PDF – the layout of OfficeMath objects is preserved automatically
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file was not created at '{pdfPath}'.");

        // Validation: ensure the document still contains top‑level OfficeMath nodes
        int topLevelCount = 0;
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelCount++;
        }

        if (topLevelCount == 0)
            throw new InvalidOperationException("No top‑level OfficeMath equations were found after conversion.");

        Console.WriteLine($"Document saved as '{docPath}' and exported to PDF '{pdfPath}'.");
        Console.WriteLine($"Top‑level OfficeMath equations count: {topLevelCount}");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and cleans up the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field (the field code "EQ" is added automatically)
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field separator – this is the documented safe pattern
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that Word evaluates the EQ code; this also prepares it for conversion
        field.Update();

        // Return the builder to the paragraph that contains the field
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node and remove the original field
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Add a blank paragraph after the equation for spacing
        builder.InsertParagraph();
    }
}
