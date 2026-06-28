using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Smart3D.ExcelImport.Tools
{
    /// <summary>
    /// Utility to generate sample Excel template files for the import operation.
    /// </summary>
    public static class SampleExcelGenerator
    {
        /// <summary>
        /// Creates a sample Excel file showing the expected format for import.
        /// </summary>
        public static void GenerateTemplate(string outputPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Import Data");

                // Headers
                worksheet.Cells[1, 1].Value = "ObjectName";
                worksheet.Cells[1, 2].Value = "ObjectType";
                worksheet.Cells[1, 3].Value = "AttributeName";
                worksheet.Cells[1, 4].Value = "AttributeValue";

                // Style headers
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }

                // Sample data rows
                var sampleData = new string[,]
                {
                    { "PIPE-RUN-001", "PipeRun", "PipingSpec", "A1A" },
                    { "PIPE-RUN-001", "PipeRun", "Service", "Cooling Water" },
                    { "PIPE-RUN-002", "PipeRun", "PipingSpec", "B2B" },
                    { "PIPE-RUN-002", "PipeRun", "OperatingTemperature", "150" },
                    { "PIPE-RUN-003", "PipeRun", "DesignPressure", "30" },
                    { "PIPELINE-100", "Pipeline", "Description", "Main Process Line" },
                    { "PIPELINE-100", "Pipeline", "Service", "Process Fluid" },
                    { "PIPELINE-200", "Pipeline", "Service", "Return Line" },
                    { "EQUIP-P01", "Equipment", "Tag", "P-101A" },
                    { "EQUIP-P01", "Equipment", "Description", "Feed Pump" },
                    { "EQUIP-T01", "Equipment", "OperatingWeight", "2500" },
                    { "EQUIP-T01", "Equipment", "Service", "Storage" },
                    { "EQUIP-E01", "Equipment", "Description", "Heat Exchanger" },
                    { "EQUIP-E01", "Equipment", "DesignWeight", "5000" },
                };

                for (int i = 0; i < sampleData.GetLength(0); i++)
                {
                    worksheet.Cells[i + 2, 1].Value = sampleData[i, 0];
                    worksheet.Cells[i + 2, 2].Value = sampleData[i, 1];
                    worksheet.Cells[i + 2, 3].Value = sampleData[i, 2];
                    worksheet.Cells[i + 2, 4].Value = sampleData[i, 3];
                }

                // Add a "Instructions" worksheet
                var instructions = package.Workbook.Worksheets.Add("Instructions");
                instructions.Cells[1, 1].Value = "Smart3D V14 - Bulk Property Import Template";
                instructions.Cells[1, 1].Style.Font.Size = 16;
                instructions.Cells[1, 1].Style.Font.Bold = true;

                instructions.Cells[3, 1].Value = "Instructions:";
                instructions.Cells[3, 1].Style.Font.Bold = true;
                instructions.Cells[4, 1].Value = "1. Fill in the 'Import Data' sheet with your object properties.";
                instructions.Cells[5, 1].Value = "2. ObjectName: The name of the Smart3D object as it appears in the model.";
                instructions.Cells[6, 1].Value = "3. ObjectType: The type - PipeRun, Pipeline, or Equipment.";
                instructions.Cells[7, 1].Value = "4. AttributeName: The property/attribute name to set.";
                instructions.Cells[8, 1].Value = "5. AttributeValue: The value to assign to the property.";
                instructions.Cells[9, 1].Value = "6. Do not modify the header row.";
                instructions.Cells[11, 1].Value = "Supported Object Types:";
                instructions.Cells[11, 1].Style.Font.Bold = true;
                instructions.Cells[12, 1].Value = "- PipeRun: Piping runs in the model";
                instructions.Cells[13, 1].Value = "- Pipeline: Pipeline systems";
                instructions.Cells[14, 1].Value = "- Equipment: Equipment objects (vessels, pumps, exchangers)";

                // Auto-fit columns
                worksheet.Columns.AutoFit();
                instructions.Columns.AutoFit();

                package.SaveAs(new FileInfo(outputPath));
            }
        }

        /// <summary>
        /// Creates a sample Excel file at the specified path.
        /// </summary>
        public static string CreateSampleFile()
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var filePath = Path.Combine(desktop, "Smart3D_ImportTemplate.xlsx");
            GenerateTemplate(filePath);
            return filePath;
        }
    }
}
