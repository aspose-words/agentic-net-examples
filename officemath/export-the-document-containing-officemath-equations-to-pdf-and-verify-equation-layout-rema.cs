using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathExportExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("OfficeMath Export Example");

        // Add a paragraph that will contain the equation.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Below is a simple fraction equation:");

        // Insert an EQ field. The field code already contains the "EQ" switch.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ field arguments (fraction 1/2) at the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // Fraction 1 over 2.

        // Update the field so that its internal state reflects the new arguments.
        eqField.Update();

        // Move back to the start of the field before conversion.
        builder.MoveTo(eqField.Start);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set display type and justification for the top‑level equation.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Add another paragraph after the equation.
        builder.Writeln();
        builder.Writeln("End of example.");

        // Verify that the document contains at least one OfficeMath node.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the document.");

        // Define output file paths.
        string outputDocx = Path.Combine(Environment.CurrentDirectory, "OfficeMathExample.docx");
        string outputPdf = Path.Combine(Environment.CurrentDirectory, "OfficeMathExample.pdf");

        // Save the document as DOCX (optional, for inspection).
        doc.Save(outputDocx, SaveFormat.Docx);

        // Save the document as PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF file was created and is not empty.
        if (!File.Exists(outputPdf) || new FileInfo(outputPdf).Length == 0)
            throw new InvalidOperationException("PDF export failed or resulted in an empty file.");

        Console.WriteLine("Document with OfficeMath equations exported to PDF successfully.");
    }
}
