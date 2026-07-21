using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Path for the output DOCX file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedOfficeMath.docx");

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Below is a sample equation created via EQ field bootstrap:");

        // Insert an empty EQ field.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write a simple fraction switch at the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Apply formatting only to top‑level OfficeMath nodes.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;
        }

        // Save the document as DOCX – this preserves the OfficeMath node and its formatting.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output DOCX file was not created.", outputPath);

        // Optional validation: ensure at least one top‑level OfficeMath node exists.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        bool hasTopLevel = false;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                hasTopLevel = true;
                break;
            }
        }

        if (!hasTopLevel)
            throw new InvalidOperationException("No top‑level OfficeMath nodes were found in the saved document.");
    }
}
