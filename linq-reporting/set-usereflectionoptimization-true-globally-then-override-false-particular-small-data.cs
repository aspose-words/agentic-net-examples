using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Set the global reflection optimization flag to true.
            //    This enables dynamic class generation for all reporting
            //    engine instances unless overridden.
            // ------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;

            // ------------------------------------------------------------
            // 2. Prepare a large data source (e.g., a list of 10,000 items).
            //    For large collections the optimization reduces overall time.
            // ------------------------------------------------------------
            var largeData = GenerateLargeData(10000);
            Document largeTemplate = new Document(); // empty template or load from file
            // (Optionally add fields to the template here.)

            // Build the report for the large data source.
            var engine = new ReportingEngine();
            engine.BuildReport(largeTemplate, largeData);
            largeTemplate.Save("LargeReport.docx");

            // ------------------------------------------------------------
            // 3. For a particular small data source we want to disable the
            //    optimization because the overhead of dynamic class generation
            //    can outweigh the benefit for tiny collections.
            // ------------------------------------------------------------
            var smallData = GenerateLargeData(5); // just a few items
            Document smallTemplate = new Document(); // empty template or load from file

            // Temporarily turn off the optimization.
            ReportingEngine.UseReflectionOptimization = false;

            // Build the report for the small data source.
            engine.BuildReport(smallTemplate, smallData);
            smallTemplate.Save("SmallReport.docx");

            // ------------------------------------------------------------
            // 4. Restore the global setting to its original value (true) so
            //    subsequent operations continue to use the optimized path.
            // ------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;
        }

        // Helper method to generate dummy data objects.
        private static List<SampleItem> GenerateLargeData(int count)
        {
            var list = new List<SampleItem>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(new SampleItem
                {
                    Id = i + 1,
                    Name = $"Item {i + 1}",
                    Description = $"Description for item {i + 1}"
                });
            }
            return list;
        }

        // Simple POCO used as a data source for the reporting engine.
        public class SampleItem
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
        }
    }
}
