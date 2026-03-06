using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating so that a template can be applied on save.
        doc.AutomaticallyUpdateStyles = true;

        // Insert an IF field that will act as a conditional block.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Result of conditional block:");
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "5";
        ifField.ComparisonOperator = "=";
        ifField.RightExpression = "5";
        ifField.TrueText = "Condition is true.";
        ifField.FalseText = "Condition is false.";
        ifField.Update();

        // Prepare SaveOptions with a default DOTX template.
        // The template will be applied because the document has no attached template.
        SaveOptions options = SaveOptions.CreateSaveOptions("ConditionalBlock.docx");
        options.DefaultTemplate = @"C:\Templates\DefaultTemplate.dotx"; // <-- path to your .dotx file

        // Save the document using the options; the default template will be attached automatically.
        doc.Save(@"C:\Output\ConditionalBlock.docx", options);
    }
}
