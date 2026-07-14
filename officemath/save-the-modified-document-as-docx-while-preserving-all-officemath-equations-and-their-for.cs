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
        // Path for the output DOCX file.
        string outputPath = "ModifiedDocument.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph.
        builder.Writeln("Sample document with OfficeMath equations:");
        builder.Writeln();

        // ------------------------------------------------------------
        // Insert the first equation using the deterministic EQ-field bootstrap.
        // ------------------------------------------------------------
        builder.Writeln("Equation 1 (fraction 1/2):");

        // Insert an empty EQ field.
        FieldEQ field1 = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ switch arguments.
        builder.MoveTo(field1.Separator);
        builder.Write(@"\f(1,2)"); // Simple fraction 1 over 2.

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field1.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath1 = field1.AsOfficeMath();
        if (officeMath1 != null)
        {
            // Insert the OfficeMath node before the field start.
            field1.Start.ParentNode.InsertBefore(officeMath1, field1.Start);
            // Remove the original field so only the OfficeMath remains.
            field1.Remove();

            // Preserve formatting: display the equation on its own line and left‑justify it.
            officeMath1.DisplayType = OfficeMathDisplayType.Display;
            officeMath1.Justification = OfficeMathJustification.Left;
        }

        // Add a blank line after the equation.
        builder.Writeln();

        // ------------------------------------------------------------
        // Insert a second equation (integral with summation) using the same pattern.
        // ------------------------------------------------------------
        builder.Writeln("Equation 2 (integral with summation):");

        FieldEQ field2 = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field2.Separator);
        builder.Write(@"\i \su(n=1,5,n)"); // Integral with summation from n=1 to 5.
        builder.MoveTo(field2.Start.ParentNode);

        OfficeMath officeMath2 = field2.AsOfficeMath();
        if (officeMath2 != null)
        {
            field2.Start.ParentNode.InsertBefore(officeMath2, field2.Start);
            field2.Remove();

            // Keep the default inline display but set justification to center.
            officeMath2.Justification = OfficeMathJustification.Center;
        }

        // Save the document as DOCX. All OfficeMath nodes retain their formatting.
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
