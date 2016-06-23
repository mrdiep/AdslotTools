using System;
using System.IO;
using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.Diagnostics;

namespace ConsoleApplication8
{
    public class CopyDataExcel
    {
        public CopyDataExcel(string directory = @"C:\GitHub\symphony\Src\ImportService.Core.UnitTests\ImportFiles\SymphonyImportTemplate\Fail - Copy")
        {
            var files = Directory.GetFiles(directory);

            foreach (var e in files)
            {
                try
                {
                    Debug.WriteLine(e);
                    var ExcelEngine1 = new ExcelEngine();
                    var Application1 = ExcelEngine1.Excel;
                    var WorkBook1 = Application1.Workbooks.Open(e, ExcelOpenType.Automatic);

                    var ExcelEngine2 = new ExcelEngine();
                    var Application2 = ExcelEngine2.Excel;
                    var WorkBook2 = Application2.Workbooks.Open(@"C:\Users\diepnguyenv\Downloads\symphonyexporttemplate_20160620 (1).xlsx", ExcelOpenType.Automatic);
                    Copy(WorkBook1, WorkBook2);
                    WorkBook2.SaveAs(@"D:\Success\" + Path.GetFileName(e));

                    WorkBook2.Close();
                    WorkBook1.Close();
                    ExcelEngine1.Dispose();
                    ExcelEngine2.Dispose();
                    Debug.WriteLine("DONE--------> " + e);
                }
                catch (Exception ex)
                {
                }
            }
        }

        private static void Copy(IWorkbook workBook1, IWorkbook workBook2)
        {
            var workSheet1 = workBook1.Worksheets[0];
            var names1 = new Dictionary<int, string>();
            var name11 = new Dictionary<string, int>();
            foreach (var t in A1)
            {
                try
                {
                    names1.Add(workSheet1[t].Column, t);
                    name11.Add(t, workSheet1[t].Column);
                }
                catch (Exception ex)
                {
                }
            }

            var names2 = new Dictionary<string, int>();
            foreach (var t in A2)
            {
                names2.Add(t, workBook2.Names[t].RefersToRange.Column);
            }

            var workSheet2 = workBook2.Worksheets[0];
            var countRow = workSheet1.Rows.Length + 1;
            var countCol = workSheet1.Columns.Length + 1;
            for (var i = 6; i <= countRow; i++)
            {
                for (var c = 1; c < workSheet1.Columns.Length; c++)
                {
                    if (workSheet1[i, c].HasFormula)
                        continue;

                    if (workSheet1[i, c].Value2 is string && string.IsNullOrEmpty(workSheet1[i, c].Value2 as string))
                        continue;

                    if (names1.ContainsKey(c))
                    {
                        var colName2 = names2[ConvertToNewName(names1[c])];
                        workSheet2[i, colName2].Value2 = workSheet1[i, c].Value2;
                    }
                }
            }

            //Copy gridDate
            var startGridDate = name11["ItemGUID"] + 1;
            var startGridDateSheet2 = names2["ItemGUID"] + 1;
            for (var rowSheet1 = 3; rowSheet1 <= countRow; rowSheet1++)
            {
                for (var colSheet1 = startGridDate; colSheet1 <= countCol; colSheet1++)
                {
                    var value2 = workSheet1[rowSheet1, colSheet1].Value2;
                    if (value2 is string && string.IsNullOrEmpty((string)value2))
                        continue;

                    var newCol = startGridDateSheet2 + (colSheet1 - startGridDate);
                    workSheet2[rowSheet1, newCol].Value2 = workSheet1[rowSheet1, colSheet1].Value2;
                }
            }
        }

        public static string ConvertToNewName(string name)
        {
            switch (name)
            {
                case "Cost9": return "VendorNetMediaCost";
                case "ClientVendorCommission": return "ClientCommission";
                case "ClientVendorCommissionType": return "ClientCommissionType";
                case "RatecardRate": return "RateCardRate";

                case "ServiceFee": return "ClientMediaServiceFee";
                case "ServiceFeeType": return "ClientMediaServiceFeeType";
                case "ServiceFeeBase": return "ClientMediaServiceFeeBase";

                case "NetMediaCost": return "VendorNetNetCost";

                case "Discount1": return "VendorDiscount1";
                case "Discount1Type": return "VendorDiscount1Type";

                case "Discount2": return "VendorDiscount2";
                case "Discount2Type": return "VendorDiscount2Type";
                case "Discount2Base": return "VendorDiscount2Base";

                case "Discount3": return "VendorDiscount3";
                case "Discount3Type": return "VendorDiscount3Type";
                case "Discount3Base": return "VendorDiscount3Base";

                case "Loading1": return "VendorLoading1";
                case "Loading1Type": return "VendorLoading1Type";
                case "Loading1Base": return "VendorLoading1Base";

                case "Loading2": return "VendorLoading2";
                case "Loading2Type": return "VendorLoading2Type";
                case "Loading2Base": return "VendorLoading2Base";

                case "Loading3": return "VendorLoading3";
                case "Loading3Type": return "VendorLoading3Type";
                case "Loading3Base": return "VendorLoading3Base";

                case "Discount4": return "VendorDiscount4";
                case "Discount4Type": return "VendorDiscount4Type";
                case "Discount4Base": return "VendorDiscount4Base";

                case "Discount5": return "VendorDiscount5";
                case "Discount5Type": return "VendorDiscount5Type";
                case "Discount5Base": return "VendorDiscount5Base";

                case "Discount6": return "VendorDiscount6";
                case "Discount6Type": return "VendorDiscount6Type";
                case "Discount6Base": return "VendorDiscount6Base";

                case "PurchaseCost": return "VendorPurchaseCost";
                case "PrepayDiscount": return "VendorPrePayDiscount";
                case "PrepayDiscountType": return "VendorPrePayDiscountType";

                case "Surcharge1": return "VendorSurcharge1";
                case "Surcharge1Type": return "VendorSurcharge1Type";
                case "Surcharge1Base": return "VendorSurcharge1Base";

                case "Surcharge2": return "VendorSurcharge2";
                case "Surcharge2Type": return "VendorSurcharge2Type";
                case "Surcharge2Base": return "VendorSurcharge2Base";

                case "Surcharge3": return "VendorSurcharge3";
                case "Surcharge3Type": return "VendorSurcharge3Type";
                case "Surcharge3Base": return "VendorSurcharge3Base";

                default:
                    return name;
            }
        }

        public static string[] A1 =
        {
          "ItemType",
            "ItemGroupLabel",
            "ItemName",
            "Publisher",
            "Site",
            "Location",
            "Format",
            "Width",
            "Height",
            "BuyType",
            "MediaType",
            "ThirdPartyTrackingType",
            "ThirdPartyTrackingTypeRate",
            "FourthPartyTrackingRate",
            "StartDate",
            "EndDate",
            "Goal",
            "ESTImpressions",
            "ESTClicks",
            "ESTAcquisitions",
            "ESTCTR",
            "ESTCVR",
            "ItemCurrency",
            "RateCardCost",
            "BaseCost",
            "RatecardRate",
            "BaseRate",
            "Discount1",
            "Discount1Type",
            "Cost1",
            "Discount2",
            "Discount2Type",
            "Discount2Base",
            "Cost2",
            "Discount3",
            "Discount3Type",
            "Discount3Base",
            "Cost3",
            "Loading1",
            "Loading1Type",
            "Loading1Base",
            "Cost4",
            "Loading2",
            "Loading2Type",
            "Loading2Base",
            "Cost5",
            "Loading3",
            "Loading3Type",
            "Loading3Base",
            "Cost6",
            "Discount4",
            "Discount4Type",
            "Discount4Base",
            "Cost7",
            "Discount5",
            "Discount5Type",
            "Discount5Base",
            "Cost8",
            "Discount6",
            "Discount6Type",
            "Discount6Base",
            "PurchaseCost",
            "VendorCommission",
            "VendorCommissionType",
            "Cost9",
            "PrepayDiscount",
            "PrepayDiscountType",
            "NetMediaCost",
            "Surcharge1",
            "Surcharge1Type",
            "Surcharge1Base",
            "Cost10",
            "Surcharge2",
            "Surcharge2Type",
            "Surcharge2Base",
            "Cost11",
            "Surcharge3",
            "Surcharge3Type",
            "Surcharge3Base",
            "AgencyTotalCost",
            "ClientBaseCost",
            "ClientBaseRate",
            "ClientDiscount1",
            "ClientDiscount1Type",
            "ClientCost1",
            "ClientDiscount2",
            "ClientDiscount2Type",
            "ClientDiscount2Base",
            "ClientCost2",
            "ClientDiscount3",
            "ClientDiscount3Type",
            "ClientDiscount3Base",
            "ClientCost3",
            "ClientLoading1",
            "ClientLoading1Type",
            "ClientLoading1Base",
            "ClientCost4",
            "ClientLoading2",
            "ClientLoading2Type",
            "ClientLoading2Base",
            "ClientCost5",
            "ClientLoading3",
            "ClientLoading3Type",
            "ClientLoading3Base",
            "ClientCost6",
            "ClientDiscount4",
            "ClientDiscount4Type",
            "ClientDiscount4Base",
            "ClientCost7",
            "ClientDiscount5",
            "ClientDiscount5Type",
            "ClientDiscount5Base",
            "ClientCost8",
            "ClientDiscount6",
            "ClientDiscount6Type",
            "ClientDiscount6Base",
            "ClientPurchaseCost",
            "ClientVendorCommission",
            "ClientVendorCommissionType",
            "ClientNetMediaCost",
            "ServiceFee",
            "ServiceFeeType",
            "ServiceFeeBase",
            "ClientCost9",
            "ClientSurcharge1",
            "ClientSurcharge1Type",
            "ClientSurcharge1Base",
            "ClientCost10",
            "ClientSurcharge2",
            "ClientSurcharge2Type",
            "ClientSurcharge2Base",
            "ClientCost11",
            "ClientSurcharge3",
            "ClientSurcharge3Type",
            "ClientSurcharge3Base",
            "ClientTotalCost",
            "BillingSource",
            "ByHour",
            "SOV",
            "IsCapped",
            "CampaignCountryBudget",
            "Classification1",
            "Classification2",
            "Classification3",
            "Comments",
            "ItemGUID",
        };

        public static string[] A2 =
        {
         "ItemType",
        "ItemGroupLabel",
        "ItemName",
        "Publisher",
        "Site",
        "Location",
        "Format",
        "Width",
        "Height",
        "BuyType",
        "MediaType",
        "ThirdPartyTrackingType",
        "ThirdPartyTrackingTypeRate",
        "FourthPartyTrackingRate",
        "StartDate",
        "EndDate",
        "Goal",
        "ESTImpressions",
        "ESTClicks",
        "ESTAcquisitions",
        "ESTCTR",
        "ESTCVR",
        "ItemCurrency",
        "RateCardCost",
        "BaseCost",
        "RateCardRate",
        "BaseRate",
        "VendorDiscount1",
        "VendorDiscount1Type",
        "Cost1",
        "VendorDiscount2",
        "VendorDiscount2Type",
        "VendorDiscount2Base",
        "Cost2",
        "VendorDiscount3",
        "VendorDiscount3Type",
        "VendorDiscount3Base",
        "Cost3",
        "VendorLoading1",
        "VendorLoading1Type",
        "VendorLoading1Base",
        "Cost4",
        "VendorLoading2",
        "VendorLoading2Type",
        "VendorLoading2Base",
        "Cost5",
        "VendorLoading3",
        "VendorLoading3Type",
        "VendorLoading3Base",
        "Cost6",
        "VendorDiscount4",
        "VendorDiscount4Type",
        "VendorDiscount4Base",
        "Cost7",
        "VendorDiscount5",
        "VendorDiscount5Type",
        "VendorDiscount5Base",
        "Cost8",
        "VendorDiscount6",
        "VendorDiscount6Type",
        "VendorDiscount6Base",
        "VendorPurchaseCost",
        "VendorCommission",
        "VendorCommissionType",
        "VendorNetMediaCost",
        "VendorPrePayDiscount",
        "VendorPrePayDiscountType",
        "VendorNetNetCost",
        "VendorSurcharge1",
        "VendorSurcharge1Type",
        "VendorSurcharge1Base",
        "Cost10",
        "VendorSurcharge2",
        "VendorSurcharge2Type",
        "VendorSurcharge2Base",
        "Cost11",
        "VendorSurcharge3",
        "VendorSurcharge3Type",
        "VendorSurcharge3Base",
        "AgencyTotalCost",
        "ClientBaseCost",
        "ClientBaseRate",
        "ClientDiscount1",
        "ClientDiscount1Type",
        "ClientCost1",
        "ClientDiscount2",
        "ClientDiscount2Type",
        "ClientDiscount2Base",
        "ClientCost2",
        "ClientDiscount3",
        "ClientDiscount3Type",
        "ClientDiscount3Base",
        "ClientCost3",
        "ClientLoading1",
        "ClientLoading1Type",
        "ClientLoading1Base",
        "ClientCost4",
        "ClientLoading2",
        "ClientLoading2Type",
        "ClientLoading2Base",
        "ClientCost5",
        "ClientLoading3",
        "ClientLoading3Type",
        "ClientLoading3Base",
        "ClientCost6",
        "ClientDiscount4",
        "ClientDiscount4Type",
        "ClientDiscount4Base",
        "ClientCost7",
        "ClientDiscount5",
        "ClientDiscount5Type",
        "ClientDiscount5Base",
        "ClientCost8",
        "ClientDiscount6",
        "ClientDiscount6Type",
        "ClientDiscount6Base",
        "ClientPurchaseCost",
        "ClientCommission",
        "ClientCommissionType",
        "ClientNetMediaCost",
        "ClientPrepayDiscount",
        "ClientPrepayDiscountType",
        "ClientNetNetCost",
        "ClientMediaServiceFee",
        "ClientMediaServiceFeeType",
        "ClientMediaServiceFeeBase",
        "ClientCost9",
        "ClientSurcharge1",
        "ClientSurcharge1Type",
        "ClientSurcharge1Base",
        "ClientCost10",
        "ClientSurcharge2",
        "ClientSurcharge2Type",
        "ClientSurcharge2Base",
        "ClientCost11",
        "ClientSurcharge3",
        "ClientSurcharge3Type",
        "ClientSurcharge3Base",
        "ClientTotalCost",
        "BillingSource",
        "ByHour",
        "SOV",
        "IsCapped",
        "CampaignCountryBudget",
        "Classification1",
        "Classification2",
        "Classification3",
        "Comments",
        "ItemGUID",
        };
    }
}
