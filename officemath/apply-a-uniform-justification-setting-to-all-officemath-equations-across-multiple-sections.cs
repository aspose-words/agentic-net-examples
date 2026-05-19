using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ApplyUniformJustification
{
    public static void Main()
    {
        // Paths for the intermediate and final documents.
        string intermediatePath = "Intermediate.docx";
        string finalPath = "Justified.docx";

        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample document with three sections, each containing two equations.
        for (int sectionIndex = 1; sectionIndex <= 3; sectionIndex++)
        {
            // Add a heading for the section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {sectionIndex}");

            // Insert two equations in the section.
            InsertOfficeMath(builder, @"\f(1,2)"); // Simple fraction 1/2
            InsertOfficeMath(builder, @"\r(3,x)"); // Cube root of x

            // Insert a section break after each section except the last.
            if (sectionIndex < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the intermediate document.
        doc.Save(intermediatePath, SaveFormat.Docx);

        // Reload the document to simulate a typical processing scenario.
        Document loadedDoc = new Document(intermediatePath);

        // Apply a uniform justification (Center) to all top‑level OfficeMath equations.
        foreach (OfficeMath om in loadedDoc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Display type must be set before justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Center;
            }
        }

        // Save the final document.
        loadedDoc.Save(finalPath, SaveFormat.Docx);

        // Validation: ensure the output file exists.
        if (!File.Exists(finalPath))
            throw new InvalidOperationException($"The file '{finalPath}' was not created.");

        // Validation: ensure every top‑level equation has the expected justification.
        Document validationDoc = new Document(finalPath);
        var mismatched = validationDoc.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .Where(om => om.MathObjectType == MathObjectType.OMathPara && om.Justification != OfficeMathJustification.Center)
            .ToList();

        if (mismatched.Any())
            throw new InvalidOperationException("One or more OfficeMath equations do not have the expected justification.");
    }

    // Helper method that creates a real OfficeMath node using the deterministic EQ‑field workflow.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is processed.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Ensure conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Move the builder back to the paragraph containing the inserted OfficeMath.
        builder.MoveTo(officeMath.ParentParagraph);
        builder.Writeln(); // Ensure subsequent content starts on a new line.
    }
}
