using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will act as the condition trigger.
        builder.Writeln("DisableAll");

        // Insert several checkbox form fields with distinct names.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.Writeln();

        // Determine if the condition is met (document contains the trigger word).
        bool conditionMet = doc.GetText().Contains("DisableAll");

        if (conditionMet)
        {
            // Disable all checkbox form fields in the document.
            foreach (FormField field in doc.Range.FormFields)
            {
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    field.Enabled = false;
                }
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
