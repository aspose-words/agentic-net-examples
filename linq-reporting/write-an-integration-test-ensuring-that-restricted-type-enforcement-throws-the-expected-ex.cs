using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedTypeTest
{
    // Sample data model.
    public class ReportModel
    {
        public SecretInfo Secret { get; set; } = new();
        public string Title { get; set; } = "Report Title";
    }

    // Type that will be restricted.
    public class SecretInfo
    {
        public string Name { get; set; } = "Classified";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words for some encodings.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Create a working directory.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Build the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that accesses a property of a restricted type.
            builder.Writeln("Secret name: <<[model.Secret.Name]>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for reporting.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the model instance.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Secret = new SecretInfo { Name = "TopSecret" },
                Title = "Demo Report"
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine with restricted types (allow only string).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The RestrictedTypes property may not be present in all versions.
            // Use reflection to set it when available.
            PropertyInfo? restrictedProp = typeof(ReportingEngine).GetProperty("RestrictedTypes",
                BindingFlags.Public | BindingFlags.Instance);
            if (restrictedProp != null && restrictedProp.CanWrite)
            {
                restrictedProp.SetValue(engine, new Type[] { typeof(string) });
            }

            // Enable inline error messages so that BuildReport returns false on failure.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // -----------------------------------------------------------------
            // 5. Attempt to build the report and verify that the expected exception is thrown.
            // -----------------------------------------------------------------
            bool exceptionThrown = false;
            try
            {
                // BuildReport returns true on success; we expect it to fail because SecretInfo is restricted.
                bool success = engine.BuildReport(loadedTemplate, model, "model");

                // If BuildReport returns false, retrieve the inline error messages.
                if (!success)
                {
                    // The engine inserts error messages into the document; we can also treat this as a failure.
                    Console.WriteLine("Test failed: BuildReport returned false without throwing an exception.");
                }
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains(nameof(SecretInfo)))
            {
                exceptionThrown = true;
                Console.WriteLine("Test passed: restricted type exception was thrown as expected.");
                Console.WriteLine($"Exception message: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Test failed: an unexpected exception was thrown.");
                Console.WriteLine($"Exception type: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
            }

            if (!exceptionThrown)
            {
                Console.WriteLine("Test failed: no exception was thrown when accessing a restricted type.");
            }

            // Optional clean‑up.
            // File.Delete(templatePath);
        }
    }
}
