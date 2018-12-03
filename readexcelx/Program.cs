using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace readexcelx
{
    class Program
    {


        static string getValue(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, int x, int y)
        {
            object value = (xlWorksheet.Cells[x, y] as Microsoft.Office.Interop.Excel.Range).get_Value(Type.Missing);
            
            return (value != null ? value.ToString().Trim() : "");
        }

        static bool setValue(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, int x, int y, string setValue)
        {
            xlWorksheet.Cells[x, y] = setValue;

            return true;
        }

        static Microsoft.Office.Interop.Excel._Worksheet setHyperLink(Microsoft.Office.Interop.Excel._Worksheet excelWorksheet, Microsoft.Office.Interop.Excel.Workbook workbk,
            int rowCount, BBDataExtractClass.Rules rule, BBDataExtractClass.policyRules pr)
        {
            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range("D" + (rowCount + 3), "D" + (rowCount + 3));
            excelWorksheet.Hyperlinks.Add(excelCell, "#rule_" + pr.PolicyID + "_" + rule.RuleID + "!A1", Type.Missing, "Client Rule", "rule_" + pr.PolicyID + "_" + rule.RuleID);

            var newSheet = workbk.Worksheets["RuleTemplate"] as Worksheet;
            newSheet.Copy(Type.Missing, workbk.Worksheets[workbk.Sheets.Count]);
            newSheet.Name = "rule_" + pr.PolicyID + "_" + rule.RuleID;
            newSheet.Cells[1048576, 1] = "rule_" + pr.PolicyID + "_" + rule.RuleID;


            Microsoft.Office.Interop.Excel.Range excelCellInner = (Microsoft.Office.Interop.Excel.Range)newSheet.get_Range("A1", "A1");
            newSheet.Hyperlinks.Add(excelCellInner, "#clientRules!A1", Type.Missing, "Client Rule", "clientRules");

            var renameSheet = workbk.Worksheets["RuleTemplate (2)"] as Worksheet;
            renameSheet.Name = "RuleTemplate";

            return newSheet;

        }

        private static void WriteToAFileWithStreamWriter(string content)
        {
            using (StreamWriter writer = File.CreateText(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] + @"\benefits_" + DateTime.Now.ToString("ddMMMyyyyhhmmss") + ".json"))
            {
                writer.WriteLine(content);
            }
        }

        public static void LoadClientRulesJsonToExcel()
        {
            using (StreamReader r = new StreamReader(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] +
                System.Configuration.ConfigurationSettings.AppSettings["clientRuleFile"]))
            {

                string json = r.ReadToEnd();
                //BBDataExtractClass.Client cll =
                //   Newtonsoft.Json.JsonConvert.DeserializeObject<BBDataExtractClass.Client>(json);

                //dynamic jsond = JObject.FromObject(cll);
                //jsond.id = cll._id;
                //var document = await Client.ReplaceDocumentAsync(cll.SelfLink, json);

                Regex rx = new Regex(@"NumberInt\((\d+)\)", RegexOptions.Multiline);
                String output = rx.Replace(json, "$1");
                try
                {
                    countryDicValue cv = new countryDicValue();
                    cv.setCountries();
                    BBDataExtractClass.ClientRules items = Newtonsoft.Json.JsonConvert.DeserializeObject<BBDataExtractClass.ClientRules>(output);
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                    //open excel file
                    Microsoft.Office.Interop.Excel.Workbook workbk = excel.Workbooks.Open(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] +
                         System.Configuration.ConfigurationSettings.AppSettings["clientRuleExcelFile"],
                      0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false,
                      System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false,
                      System.Reflection.Missing.Value, false, false, false);

                    Microsoft.Office.Interop.Excel.Sheets xlsheets = workbk.Sheets;
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlsheets["clientRules"];

                    setValue(excelWorksheet, 3, 1, items.ClientNo);
                    int rowCount = 0;


                    foreach (BBDataExtractClass.policyRules pr in items.Policies)
                    {
                        Console.WriteLine(string.Format("Policy ID: {0}, Policy Name : {1}", pr.PolicyID, pr.PolicyName));

                        setValue(excelWorksheet, 3 + rowCount, 2, pr.PolicyID.ToString());
                        setValue(excelWorksheet, 3 + rowCount, 3, pr.PolicyName);

                        foreach (BBDataExtractClass.Rules rule in pr.Rules)
                        {
                            int benefitRuleCount = 0;
                            var newSheet = setHyperLink(excelWorksheet, workbk, rowCount, rule, pr);
                            setValue(newSheet, 3 + benefitRuleCount, 1, rule.TotalPoints.ToString());
                            setValue(newSheet, 3 + benefitRuleCount, 2, rule.RuleID.ToString());
                            setValue(newSheet, 3 + benefitRuleCount, 8, (rule.SpouseMoving != null ? rule.SpouseMoving.ToUpperInvariant() : ""));
                            if (rule.DepartureCountryCd != null)
                            {
                                List<string> selectedCountry = new List<string>();
                                foreach (string str in rule.DepartureCountryCd)
                                {
                                    if (cv.getCountries().ContainsKey(str))
                                    {
                                        selectedCountry.Add(cv.getCountries()[str]);
                                    }
                                }
                                setValue(newSheet, 3 + benefitRuleCount, 9, string.Join(";", selectedCountry));
                            }
                            if (rule.DestinationCountryCd != null)
                            {
                                List<string> selectedCountry = new List<string>();
                                foreach (string str in rule.DestinationCountryCd)
                                {
                                    if (cv.getCountries().ContainsKey(str))
                                    {
                                        selectedCountry.Add(cv.getCountries()[str]);
                                    }
                                }

                                setValue(newSheet, 3 + benefitRuleCount, 10, string.Join(";", selectedCountry));
                            }
                            setValue(newSheet, 3 + benefitRuleCount, 11, (rule.JobCode != null ? string.Join(";", rule.JobCode) : ""));
                            setValue(newSheet, 3 + benefitRuleCount, 12, rule.FamilyMoving);
                            setValue(newSheet, 3 + benefitRuleCount, 13, rule.IntraRegionalFlg.ToString().ToUpperInvariant());

                            foreach (BBDataExtractClass.BenefitsRules bf in rule.Benefits)
                            {
                                setValue(newSheet, 3 + benefitRuleCount, 3, bf.BenefitID);
                                setValue(newSheet, 3 + benefitRuleCount, 4, bf.Points?.ToString());
                                setValue(newSheet, 3 + benefitRuleCount, 5, bf.Attribute);
                                setValue(newSheet, 3 + benefitRuleCount, 6, bf.IsRecommended.ToString().ToUpperInvariant());
                                setValue(newSheet, 3 + benefitRuleCount, 7, bf.ServiceType.ToUpperInvariant());
                                benefitRuleCount++;
                                Console.WriteLine(string.Format("Policy : {0}, Rule ID : {1}, Benefit ID : {2}", pr.PolicyID, rule.RuleID, bf.BenefitID));
                            }

                            rowCount++;

                        }
                        rowCount++;
                    }

                    workbk.Save(); // save 
                    workbk.Close();
                    excel.Quit();// close. dont forget this

                }
                catch (Exception ex)
                {

                    // throw;
                }

            }
        }

        public static void LoadClientsJsonToExcel()
        {
            using (StreamReader r = new StreamReader(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] +
                System.Configuration.ConfigurationSettings.AppSettings["clientFile"]))
            {

                string json = r.ReadToEnd();
                //BBDataExtractClass.Client cll =
                //   Newtonsoft.Json.JsonConvert.DeserializeObject<BBDataExtractClass.Client>(json);

                //dynamic jsond = JObject.FromObject(cll);
                //jsond.id = cll._id;
                //var document = await Client.ReplaceDocumentAsync(cll.SelfLink, json);

                Regex rx = new Regex(@"NumberInt\((\d+)\)", RegexOptions.Multiline);
                String output = rx.Replace(json, "$1");
                try
                {
                    BBDataExtractClass.Client items = Newtonsoft.Json.JsonConvert.DeserializeObject<BBDataExtractClass.Client>(output);

                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.DisplayAlerts = false;
                    //open excel file


                    foreach (BBDataExtractClass.Policies pr in items.Policies)
                    {
                        Console.WriteLine(string.Format("Policy ID: {0}, Policy Name : {1}", pr.PolicyID, pr.PolicyName));
                        Microsoft.Office.Interop.Excel.Workbook workbk = excel.Workbooks.Open(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] +
                         System.Configuration.ConfigurationSettings.AppSettings["clientExcelFile"],
                           0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false,
                           System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false,
                           System.Reflection.Missing.Value, false, false, false);

                        Microsoft.Office.Interop.Excel.Sheets xlsheets = workbk.Sheets;
                        Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlsheets["Homeowner"];

                        //setValue(excelWorksheet, 3, 1, items.ClientNo);

                        setValue(excelWorksheet, 2, 2, pr.PolicyID.ToString());
                        setValue(excelWorksheet, 1, 2, pr.PolicyName);
                        int rowCount = 5;


                        foreach (BBDataExtractClass.BenefitCards bc in pr.Benefits)
                        {

                            setValue(excelWorksheet, rowCount, 1, bc.BenefitID.ToString());
                            setValue(excelWorksheet, rowCount, 2, bc.ImageURL.ToString());
                            setValue(excelWorksheet, rowCount, 3, "Single Selection");
                            setValue(excelWorksheet, rowCount, 4, bc.Points.ToString());
                            setValue(excelWorksheet, rowCount, 5, bc.ServiceType.ToString());
                            setValue(excelWorksheet, rowCount, 6, bc.Category.ToString());

                            setValue(excelWorksheet, rowCount, 8, bc.ClientBenefitTitle);
                            setValue(excelWorksheet, rowCount, 9, bc.cardSequence.ToString());
                            setValue(excelWorksheet, rowCount, 10, bc.ClientBenefitDesc.ToString());

                            setValue(excelWorksheet, rowCount, 11, string.Join(";", bc.OrBenefits).ToString());
                            setValue(excelWorksheet, rowCount, 13, string.Join(";", bc.AndBenefits).ToString());
                            setValue(excelWorksheet, rowCount, 31, bc.ProdName.ToString());
                            setValue(excelWorksheet, rowCount, 32, bc.SubProdName.ToString());


                            rowCount++;

                        }



                        excel.DisplayAlerts = false;

                        workbk.SaveAs(System.Configuration.ConfigurationSettings.AppSettings["JsonPath"] + @"\Client_" + items.ClientNo + "_Policy_" + pr.PolicyID + ".xlsm",
                            Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange,
                                        XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                        // save 
                        workbk.Close();
                        excel.Quit();// close. dont forget this
                    }

                }
                catch (Exception ex)
                {

                    // throw;
                }

            }
        }

        public static void LoadExcelToClientsJson()
        {
            BBDataExtractClass.Client cl = new BBDataExtractClass.Client();
            cl.Policies = new List<BBDataExtractClass.Policies>();
            BBDataExtractClass.Config con = new BBDataExtractClass.Config();

            con.BenefitOrder = new List<string>();
            con.CategorySortOrder = new List<string>();
            string collectionID = System.Configuration.ConfigurationSettings.AppSettings["collectionID"];
            cl._id = @"ObjectId(""test"")";

            foreach (string excelFile in System.Configuration.ConfigurationSettings.AppSettings["excelFile"].Split(';'))
            {
                if (string.IsNullOrEmpty(excelFile))
                    continue;

                try
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(System.Configuration.ConfigurationSettings.AppSettings["excelPath"]
                        + excelFile, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    List<BBDataExtractClass.BenefitCards> benLst = new List<BBDataExtractClass.BenefitCards>();
                    BBDataExtractClass.Policies _pol = new BBDataExtractClass.Policies();

                    _pol.PolicyID = (getValue(xlWorksheet, 2, 2));
                    _pol.PolicyName = getValue(xlWorksheet, 1, 2);


                    Console.WriteLine(string.Format("Policy id: {0}", getValue(xlWorksheet, 2, 2)));
                    Console.WriteLine(string.Format("Policy Name: {0}", getValue(xlWorksheet, 1, 2)));
                    Console.WriteLine(string.Format("Policy Points: {0}", getValue(xlWorksheet, 2, 6)));

                    for (int i = 5; i <= 125; i++)
                    {
                        if (String.IsNullOrEmpty(getValue(xlWorksheet, i, 1)))
                            continue;

                        if ("ZZZCash Out" == getValue(xlWorksheet, i, 1))
                            Console.Write("wait");
                        if ("5731498185E" == getValue(xlWorksheet, i, 1))
                            Console.Write("");

                        Console.WriteLine(string.Format("Benefit ID: {0}", getValue(xlWorksheet, i, 1)));
                        Console.WriteLine(string.Format("Benefit Card Image: {0}", getValue(xlWorksheet, i, 2)));
                        Console.WriteLine(string.Format("Default value: {0}", getValue(xlWorksheet, i, 4)));
                        Console.WriteLine(string.Format("Service Type: {0}", getValue(xlWorksheet, i, 5)));
                        Console.WriteLine(string.Format("Card category: {0}", getValue(xlWorksheet, i, 6)));
                        Console.WriteLine(string.Format("Card order: {0}", getValue(xlWorksheet, i, 7)));
                        Console.WriteLine(string.Format("Card Title: {0}", getValue(xlWorksheet, i, 8)));
                        Console.WriteLine(string.Format("Card Sequence: {0}", getValue(xlWorksheet, i, 9)));
                        Console.WriteLine(string.Format("Card text/PDF content: {0}", getValue(xlWorksheet, i, 10)));

                        Console.WriteLine(string.Format("Product Name: {0}", getValue(xlWorksheet, i, 20)));
                        Console.WriteLine(string.Format("Sub Product Name: {0}", getValue(xlWorksheet, i, 21)));
                        List<string> orBenefits = null;

                        if (!String.IsNullOrEmpty(getValue(xlWorksheet, i, 11)) && getValue(xlWorksheet, i, 11).ToLowerInvariant() != "na")
                        {
                            orBenefits = new List<string>();
                            orBenefits.AddRange(getValue(xlWorksheet, i, 11).Split(';').ToList());
                        }
                        List<string> andBenefits = null;
                        if (!String.IsNullOrEmpty(getValue(xlWorksheet, i, 13)) && getValue(xlWorksheet, i, 13).ToLowerInvariant() != "na")
                        {
                            andBenefits = new List<string>();
                            andBenefits.AddRange(getValue(xlWorksheet, i, 13).Split(';').ToList());
                        }


                        benLst.Add(new BBDataExtractClass.BenefitCards()
                        {

                            BenefitID = getValue(xlWorksheet, i, 1),
                            Attribute = "",
                            CashOutValue = Math.Round(0.0, 0),
                            Category = getValue(xlWorksheet, i, 6),
                            ServiceType = getValue(xlWorksheet, i, 5),
                            ClientBenefitDesc = getValue(xlWorksheet, i, 10),
                            ClientBenefitTitle = getValue(xlWorksheet, i, 8),
                            ClientNo = getValue(xlWorksheet, i, 1).Substring(0, 4),
                            HasQuantity = false,
                            ImageURL = getValue(xlWorksheet, i, 2) + (getValue(xlWorksheet, i, 2).IndexOf(".") > 1 ? "" : ".png"),
                            IsRecommended = false,
                            LastUpdatedBy = "",
                            OrBenefits = orBenefits,
                            AndBenefits = andBenefits,
                            Points = getValue(xlWorksheet, i, 4),
                            ProdName = getValue(xlWorksheet, i, 31),
                            SubProdName = getValue(xlWorksheet, i, 32),
                            ProductNo = "0",
                            SubProductNo = "0",
                            cardSequence = Int32.Parse(getValue(xlWorksheet, i, 9)),
                            ConsultantOnly = (getValue(xlWorksheet, i, 25).ToLowerInvariant() == "true" ? true : false),
                            Hide = (getValue(xlWorksheet, i, 25).ToLowerInvariant() == "true" ? true : false)

                        });

                        Console.WriteLine("\n\nNext \n\n");
                    }



                    ///Add Cash out benefit
                    ///
                    //benLst.Add(new BBDataExtractClass.BenefitCards()
                    //{

                    //    BenefitID = benLst.FirstOrDefault().ClientNo + "99999999",
                    //    Attribute = "",
                    //    CashOutValue = 0,
                    //    Category = "Cash Out",
                    //    ClientBenefitDesc = "You have the option to convert some or all of your Flex points to cash. While a full cash-out of all points is generally not recommended, this offers the flexibility to gain access to additional discretionary funds.",
                    //    ClientBenefitTitle = "Trade Points for Cash",
                    //    ClientNo = benLst.FirstOrDefault().ClientNo,
                    //    HasQuantity = false,
                    //    ImageURL = "33.png",
                    //    IsRecommended = false,
                    //    LastUpdatedBy = "",
                    //    OrBenefits = new List<string>(),
                    //    AndBenefits = new List<string>(),
                    //    Points = "0",
                    //    ProdName = "0",
                    //    ProductNo = "0",
                    //    SubProductNo = "0",
                    //    cardSequence = benLst.Count

                    //});


                    benLst = benLst.OrderBy(x => x.cardSequence).ToList();
                    _pol.Benefits = benLst;
                    cl.ClientNo = benLst.FirstOrDefault().ClientNo;

                    con.DisableBeforeWelcomeCall = true;

                    con.BenefitOrder.AddRange(benLst.Select(i => i.BenefitID).ToList());
                    con.CategorySortOrder.AddRange(benLst.Select(i => i.Category).Distinct().ToList());



                    _pol.CashOut = new BBDataExtractClass.CashOut()
                    {
                        BasedOn = new List<string>() { "PointsCashedOut" },
                        CashOutRules = new List<BBDataExtractClass.CashOutRules>() {
                        new BBDataExtractClass.CashOutRules() {
                      //  Amount = 1000,
                      //  Currency = "USD",
                        JobCode = new List<int>() ,
                        PointsCashedOut = new List<double>() //{ 1,2,3,4,5,6,7,8}
                            }
                        }
                    };

                    cl.Config = con;
                    cl.Policies.Add(_pol);
                    _pol.PDFSalutation = System.Configuration.ConfigurationSettings.AppSettings["PDFSalutation"];

                    xlWorkbook.Close();
                    xlApp.Quit();
                }
                catch (Exception ex)
                {


                }
            }

            cl.Config.BenefitOrder = cl.Config.BenefitOrder.Distinct().ToList();

            cl.Config.CategorySortOrder = cl.Config.CategorySortOrder.Distinct().ToList();
            string outputStr = Newtonsoft.Json.JsonConvert.SerializeObject(cl,
                Newtonsoft.Json.Formatting.Indented,
                new DefaultJsonSerializer());
            outputStr = outputStr.Replace("{\"_id\":\"ObjectId", "{\"_id\":ObjectId");
            outputStr = outputStr.Replace("\\\"test\\\")\",", '"' + collectionID + '"' + "),");
            WriteToAFileWithStreamWriter(outputStr);
        }

        static void Main(string[] args)
        {
            ///1. Extracting data from Clients Rules JSON file to pre defined Spread sheet
            //  LoadClientRulesJsonToExcel();

            ///2. Extracting data from Clients JSON file to pre defined Spread sheet
            //  LoadClientsJsonToExcel();

            ///3. Exporting data from Spread sheet containing changes from SME's to Clients Rules JSON file 
            LoadExcelToClientsJson();

        }

    }

}

