using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using TheTool;

namespace ValidationTool
{
    class Program
    {
        private TableSetting[] tableSettings;
        private BaseConfig baseConfig;
        private XLSXToSQLMultiSingleConfig config;
        private Dictionary<string, string> tableNameMaps = new Dictionary<string, string>();

        private string outputDirectory;
        
        static void Main(string[] args)
        {
            File.WriteAllText("log.txt", string.Empty);

            Program program = new Program();

            program.CreateOutputDirectory();

            Console.WriteLine("Extracting sheets...");
            program.ExtractMasterSheets();
            Console.WriteLine("Validating...");
            program.Validate();
            Console.WriteLine("Cleaning up...");
            program.Cleanup();
        }

        public void CreateOutputDirectory()
        {
            string validatedPath = Path.Combine(
                    Directory.GetCurrentDirectory(),
                    @"output\validated\");

            int index = 0;
            string[] files = Directory.GetDirectories(validatedPath);
            if (files.Length > 0)
            {
                foreach (string file in files)
                {
                    int number = int.Parse(Path.GetFileName(file).Split('.')[0]);
                    if (number > index) index = number;
                }
            }
            index++;
            
            outputDirectory = Path.Combine(validatedPath, DateTime.Now.ToString(index + ". yyyyMMdd_hhmm_tt"));
            Directory.CreateDirectory(outputDirectory);
        }

        public void WriteTableSettingsConfig()
        {
            List<TableSetting> tableSettings = new List<TableSetting>();
        }

        public Program()
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            baseConfig = serializer.Deserialize<BaseConfig>(File.ReadAllText("config\\baseConfig.json"));
            config = serializer.Deserialize<XLSXToSQLMultiSingleConfig>(File.ReadAllText("config\\xlsxToSqlMultiSingleConfig.json"));
            List<TableSetting> tableSettings = serializer.Deserialize<List<TableSetting>>(File.ReadAllText("config\\tablesConfig.json"));
            this.tableSettings = tableSettings.ToArray();
            foreach (TableSetting tableSetting in tableSettings)
                tableNameMaps.Add(tableSetting.Name, ValidateTableName(tableSetting.Name));
        }

        public void Cleanup()
        {
            string[] files = Directory.GetFiles(
                Path.Combine(
                    Directory.GetCurrentDirectory(),
                    @"output\extracted"));
            foreach (string file in files)
                File.Delete(file);
        }

        public void Log(string text)
        {
            Console.WriteLine(text);

            File.AppendAllText("log.txt", text);
            File.AppendAllText("log.txt", Environment.NewLine);
        }

        public void Validate()
        {
            foreach (KeyValuePair<string, string> pair in tableNameMaps)
            {
                try
                {
                    TableSetting setting = GetSetting(pair.Key);
                    if (setting.PrimaryKeyColumns == null || setting.PrimaryKeyColumns == string.Empty)
                        continue;

                    Log("Validating " + pair.Key);

                    ComparisonData data = new ComparisonData();
                    data.keyColumns = setting.PrimaryKeyColumns.Split(',');
                    data.additionalKeyColumns = setting.SecondaryKeyColumns.Split(',');
                    data.manualCalculation = setting.ManualCalculation;
                    data.valuesToNeglect = baseConfig.ValuesToNeglect;
                    data.substringsToNeglect = baseConfig.SubstringsToNeglect;

                    if (pair.Value == "vAll")
                    {
                        data.ignoredColumns = "A,D,K,N,S".Split(',');
                        Console.WriteLine("Ignored columns for Calculations");
                    }


                    RoundingSettings settings = new RoundingSettings();
                    settings.smartRounding = baseConfig.SmartRounding;
                    settings.smartRoundingFormatting = baseConfig.SmartRoundingFormatting;
                    settings.delta = baseConfig.SmartRoundingDelta;
                    

                    XLSXSQLComparer comparer = new XLSXSQLComparer(
                        baseConfig.ValidationToolConnectionString,
                        config.ConnectionString,
                        string.Format("SELECT * FROM [{0}].[{1}]", config.Schema, pair.Value),
                        Path.Combine(
                            Directory.GetCurrentDirectory(),
                            string.Format(@"output\extracted\{0}.xlsx", pair.Value)
                            ),
                        data,
                        settings
                    );
                    XLSXDiscrepancyOutputter outputter = new XLSXDiscrepancyOutputter(
                        Path.Combine(
                            outputDirectory,
                            string.Format(@"{0}.xlsx", pair.Value)
                            ));
                    ComparisonResult result;
                    comparer.Compare(outputter, out result);
                    comparer.Cleanup();
                    outputter.Cleanup();
                } catch (Exception e)
                {
                    Log("Error: " + e.Message);
                } 
            }
            
        }

        public void ExtractMasterSheets()
        {
            Application application = new Application();
            application.ScreenUpdating = false;
            application.DisplayAlerts = false;

            Workbook workbook = application.Workbooks.Open(
                config.MasterFilePath);

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (GetSetting(worksheet.Name) == null) continue;

                Console.WriteLine("Extracting " + worksheet.Name);

                // Create new workbook.
                Workbook newWorkbook = application.Workbooks.Add();

                int dataStartRow = GetSetting(worksheet.Name).HeaderRow;

                worksheet.Range[
                    worksheet.Range["A" + dataStartRow],
                    worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)]
                .Copy(
                newWorkbook.Worksheets[1].Range["A1"]);

                string newWorksheetName = ValidateTableName(worksheet.Name);

                newWorkbook.SaveAs(Path.Combine(
                        Directory.GetCurrentDirectory(),
                        string.Format(@"output\extracted\{0}.xlsx", newWorksheetName)
                        ));
                newWorkbook.Close();
                Marshal.ReleaseComObject(newWorkbook);
            }

            workbook.Close();
            Marshal.ReleaseComObject(workbook);

            application.Quit();
            Marshal.ReleaseComObject(application);
        }

        private string ValidateTableName(string tableName)
        {
            if (tableName.IndexOf(config.Prefix) < 0)
                tableName = config.Prefix + tableName;

            foreach (SubstringMap substringMap in config.SubstringMaps)
                tableName = tableName.Replace(substringMap.From, substringMap.To);

            if (tableName.IndexOf(config.Suffix) < 0)
                tableName = tableName + config.Suffix;

            return tableName;
        }

        private TableSetting GetSetting(string tableName)
        {
            foreach (TableSetting tableSetting in tableSettings)
                if (tableSetting.Name == tableName)
                    return tableSetting;

            return null;
        }
    }
}
