using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Data.SqlClient;
using System.Drawing;
using System.Text.RegularExpressions;
using System.IO.Compression;
using ClosedXML.Excel;
using OfficeOpenXml;  //EPPLUS
using XLS = OfficeOpenXml.Style;
using Spire.Xls;



namespace BB8_Eevee_NB
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;  //LicenseContext.Commercial
            SqlConnection conn = new SqlConnection("Data Source=G8W08746A; Initial Catalog=PSGO_PBM; Persist Security Info=True; User ID=dev_project; Password=power@HP1234");
            conn.Open();

            int weekday = (int)DateTime.UtcNow.DayOfWeek;
            string ver = DateTime.UtcNow.ToString("yyyyMMdd") + "T" + DateTime.UtcNow.ToString("HHmmss");
            string month = DateTime.UtcNow.ToString("yyyyMM01");
            string[] Extend6U = new string[3] { "R6U", "RPL6U", "RT6U" };

        
            string[] sharepoint = null;
            if (weekday == 2)
            {
                sharepoint = new string[9]{
                    @"\\Raichu\Users\tangk\HP Inc\Inventec Notebook Quotes - Current CPC Trackers",
                    @"\\Raichu\Users\tangk\HP Inc\Compal Notebook Quotes - Current CPC Tracker",
                    @"\\Raichu\Users\tangk\HP Inc\Quanta Notebook Quotes - Current CPCT",
                    @"\\Raichu\Users\tangk\HP Inc\Wistron Quote SharePoint Site - Current CPCT",
                    @"\\Raichu\Users\tangk\HP Inc\Huaqin NB Quotes - Current CPCT",
                    @"\\Raichu\Users\tangk\HP Inc\Pegatron Commercial Desktop - Current CPCT",
                    @"\\Raichu\Users\tangk\HP Inc\BYD Quotes Site - Current CPC Trackers",
                    @"\\Raichu\Users\tangk\HP Inc\Team Site - Current CPC Tracker-bNB",
                    @"\\Raichu\Users\tangk\HP Inc\Team Site - Shared Documents\Current CPC Trackers"
                };
            }
            else if (weekday == 4)
            {
                sharepoint = new string[2]{
                    @"\\Raichu\Users\tangk\HP Inc\Team Site - Current CPC Tracker-bNB",
                    @"\\Raichu\Users\tangk\HP Inc\Team Site - Shared Documents\Current CPC Trackers"
                };
            }
          

            //string[] sharepoint = new string[1] { @"\\Raichu\Users\tangk\HP Inc\psgopbm-sharing - Documents\Test" };

            //string[] sharepoint = new string[9]{
            //    @"\\Raichu\Users\tangk\HP Inc\Inventec Notebook Quotes - Current CPC Trackers",
            //    @"\\Raichu\Users\tangk\HP Inc\Compal Notebook Quotes - Current CPC Tracker",
            //    @"\\Raichu\Users\tangk\HP Inc\Quanta Notebook Quotes - Current CPCT",
            //    @"\\Raichu\Users\tangk\HP Inc\Wistron Quote SharePoint Site - Current CPCT",
            //    @"\\Raichu\Users\tangk\HP Inc\Huaqin NB Quotes - Current CPCT",
            //    @"\\Raichu\Users\tangk\HP Inc\Pegatron Commercial Desktop - Current CPCT",
            //    @"\\Raichu\Users\tangk\HP Inc\BYD Quotes Site - Current CPC Trackers",
            //    @"\\Raichu\Users\tangk\HP Inc\Team Site - Current CPC Tracker-bNB",
            //    @"\\Raichu\Users\tangk\HP Inc\Team Site - Shared Documents\Current CPC Trackers"
            //};

            string[] outcome = new string[10]{
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\NB Upload Template - iCost",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\NB Upload Template - iCost NBFA\CCS NBFA - EffDate " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\NB Upload Template - MS4 PIR Creation\NB PIR Creation - " + ver + ".csv",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\NB Upload Template - MS4 PIR Price Condition\NB PIR Price Condition - " + ver + ".csv",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - Option Adder\NB Option Adder Report - " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - BUSA\NB BUSA Report - " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - BB8 Adders\NB BB8 Adders Report - " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - Cost Difference\NB Cost Difference Report - " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - Negative Cost\NB Negative Cost Report - " + ver + ".xlsx",
                @"\\Raichu\users\tangk\HP Inc\psgopbm-sharing - Documents\EeveeTool\Reporting - ODM Managed Cost\BUSA ODM Managed Cost - " + month + ".xlsx"
            };

            DataTable CPCT = ReadCPCT(conn, ver, sharepoint);

            //QTool(conn, CPCT, ver);



            DataTable dtBUSA = BUSA(CPCT, ver);
            DataTable dtHierarchy = UploadBUSA(conn, CPCT, dtBUSA, ver);

            var Options = Option(CPCT, dtHierarchy, Extend6U, ver);
            DataTable dtAdder = Options.Item1;
            DataTable dtRCTO = Options.Item2;
            DataTable dtRCTO_PN = Options.Item3;
            DataTable dtConfig_PN = Options.Item4;
            DataTable dtExclude = Options.Item5;

            UploadOption(conn, CPCT, dtAdder, dtRCTO, dtRCTO_PN, dtConfig_PN, dtExclude, weekday, ver);
            DataTable dtOwner = WriteResult(conn, CPCT, dtHierarchy, dtExclude, ver);
            List<string> allOwner = SendEmail(CPCT, dtOwner);

            string[] validation = Report(conn, CPCT, weekday, ver, outcome);
            SendEmail_icost(conn, CPCT, ver, validation, outcome, allOwner);
        }

        static DataTable ReadCPCT(SqlConnection conn, string ver, string[] sharepoint)
        {
            DataTable CPCT = new DataTable();
            CPCT.Columns.Add("filename");
            CPCT.Columns.Add("trackingname");
            CPCT.Columns.Add("filepath");
            CPCT.Columns.Add("UploadTime");
            CPCT.Columns.Add("Owner");
            CPCT.Columns.Add("Platform");
            CPCT.Columns.Add("ODM");
            CPCT.Columns.Add("RCTO_flag");
            CPCT.Columns.Add("RCTO_PN_flag");
            CPCT.Columns.Add("Null_flag");
            CPCT.Columns.Add("Error_flag");
            CPCT.Columns.Add("ErrorMsg");
            CPCT.Columns.Add("MultipleODM");
            CPCT.Columns.Add("version");
            CPCT.Columns.Add("CurrentMonth");
            CPCT.Columns["version"].DefaultValue = ver;
            CPCT.Columns["CurrentMonth"].DefaultValue = DateTime.UtcNow.ToString("yyyy-MM-01");

            for (int X = 0; X < sharepoint.Length; X++)
            {
                string[] collection = Directory.GetFiles(sharepoint[X], "*.xlsx");
                for (int Y = 0; Y < collection.Length; Y++)
                {
                    FileInfo fi = new FileInfo(collection[Y]);
                    if (!fi.Name.StartsWith("~$"))
                    {
                        string tracking = "";
                        if (fi.Name.Contains("}"))     //Regional site: TH, MX
                        {
                            string[] track_1 = fi.Name.Split(new string[] { "}" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_2 = track_1[1].Split(new string[] { "CPC" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_3 = track_2[0].Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_4 = track_3[0].Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_5 = track_4[0].Split(new string[] { "(" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_6 = track_5[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            tracking = track_6[0];
                        }
                        else
                        {
                            string[] track_2 = fi.Name.Split(new string[] { "CPC" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_3 = track_2[0].Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_4 = track_3[0].Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_5 = track_4[0].Split(new string[] { "(" }, StringSplitOptions.RemoveEmptyEntries);
                            string[] track_6 = track_5[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            tracking = track_6[0];   
                        }

                        if (tracking.Count(char.IsDigit) == 0)
                        {
                            tracking = tracking.Replace(".", string.Empty);
                        }
                        if (tracking.Length > 4 && tracking.Count(char.IsDigit) >= 2)    //If length <= 5 (B52), 6U
                        {
                            tracking = Regex.Replace(tracking, @"[\d-]", string.Empty).Replace(".", string.Empty);  //for 1.0 2.0
                        }
                        if (fi.Name.Contains("6U") && !tracking.Contains("6U"))  //Check "6U"
                        {
                            tracking = tracking + "6U";
                        }

                        DataRow row = CPCT.NewRow();
                        row["filename"] = fi.Name;
                        row["trackingname"] = tracking;
                        row["filepath"] = fi.FullName;
                        row["UploadTime"] = fi.LastWriteTimeUtc;
                        CPCT.Rows.Add(row);
                    }
                }
            }
            return CPCT;
        }

        static DataTable BUSA(DataTable CPCT, string ver)
        {
            DataTable dtBUSA = new DataTable();
            dtBUSA.Columns.Add("BUSA");
            dtBUSA.Columns.Add("Description");
            dtBUSA.Columns.Add("Total_Cost");
            dtBUSA.Columns.Add("ODM_Managed_Cost");
            dtBUSA.Columns.Add("VA");
            dtBUSA.Columns.Add("Date_Added");
            dtBUSA.Columns.Add("Platform");
            dtBUSA.Columns.Add("ODM");
            dtBUSA.Columns.Add("Comment");
            dtBUSA.Columns.Add("Flag");
            dtBUSA.Columns.Add("UploadTime");
            dtBUSA.Columns.Add("filepath");
            dtBUSA.Columns.Add("version");
            dtBUSA.Columns["ODM_Managed_Cost"].DefaultValue = 0;
            dtBUSA.Columns["VA"].DefaultValue = 0;
            dtBUSA.Columns["Flag"].DefaultValue = "NB";
            dtBUSA.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtBUSA.Columns["version"].DefaultValue = ver;

            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                string filename = CPCT.Rows[i]["filename"].ToString().Trim();
                string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                string tracking = CPCT.Rows[i]["trackingname"].ToString().Trim();

                using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet xlSheet = ep.Workbook.Worksheets["BU SA"];
                        if (xlSheet != null)
                        {
                            //01: Check headings (*check "PassThru" for Crux)
                            int passthru = 0;
                            string[] colname_1 = new string[] { "Level 3", "Description", "Base Unit Cost", "ODM Managed Cost", "VA", "Date", "Platform", "ODM" };
                            string[] colname_2 = new string[] { "Level 3", "Description", "Base Unit Cost", "ODM Managed Cost", "VA", "Date", "Platform", "Pass", "ODM" };

                            for (int X = 0; X < colname_1.Length; X++)
                            {
                                if (!xlSheet.Cells[1, 1 + X].Text.ToString().Contains(colname_1[X]))
                                {
                                    for (int Y = 0; Y < colname_2.Length; Y++)
                                    {
                                        passthru = 1;
                                        if (!xlSheet.Cells[1, 1 + Y].Text.ToString().Contains(colname_2[Y]))
                                        {
                                            CPCT.Rows[i]["Error_flag"] = 1;
                                            CPCT.Rows[i]["ErrorMsg"] = "The heading of BUSA sheet does not align with template.";
                                            goto End;
                                        }
                                    }
                                }
                            }

                            //02: Find Platform ODM (*replace IEC with Inventec)(*Multi ODM flag)(*allow 5 empty rows in between)
                            int rownum = 0;
                            string strPlatform = "", strODM = "";
                            for (int X = 2; X <= xlSheet.Dimension.End.Row; X++)
                            {
                                string platform = xlSheet.Cells[X, 7].Text.ToString() == "" ? "" : xlSheet.Cells[X, 7].Text.ToString().Trim();
                                string odm = xlSheet.Cells[X, 8 + passthru].Text.ToString() == "" ? "" : xlSheet.Cells[X, 8 + passthru].Text.ToString().Replace("IEC", "Inventec").Replace("CQ", "").Trim();  //.Replace("IMX", "MX")

                                if (platform != "" && odm != "")
                                {
                                    if (strPlatform == "" && strODM == "")
                                    {
                                        strPlatform = platform;
                                        strODM = odm;
                                    }
                                    else
                                    {
                                        if (platform.Length < strPlatform.Length)
                                        {
                                            strPlatform = platform;
                                        }
                                        if (odm != strODM)
                                        {
                                            strODM = "";
                                            CPCT.Rows[i]["MultipleODM"] = 1;
                                        }
                                    }

                                    if (xlSheet.Cells[X, 1].Text.ToString() != "" && xlSheet.Cells[X, 3].Value != null && Regex.IsMatch(xlSheet.Cells[X, 3].Value?.ToString(), @"^-?\d+"))  //+- digital number (more)
                                    {
                                        DataRow row = dtBUSA.NewRow();
                                        row["BUSA"] = xlSheet.Cells[X, 1].Text.ToString().Trim();
                                        row["Description"] = xlSheet.Cells[X, 2].Text.ToString().Trim();
                                        row["Total_Cost"] = Convert.ToDouble(xlSheet.Cells[X, 3].Value.ToString().Trim());

                                        if (xlSheet.Cells[X, 4].Text.ToString() != "" && Regex.IsMatch(xlSheet.Cells[X, 4].Value.ToString(), @"^-?\d+"))
                                        {
                                            row["ODM_Managed_Cost"] = Convert.ToDouble(xlSheet.Cells[X, 4].Value.ToString().Trim());
                                        }
                                        if (xlSheet.Cells[X, 5].Text.ToString() != "" && Regex.IsMatch(xlSheet.Cells[X, 5].Value.ToString(), @"^-?\d+"))
                                        {
                                            row["VA"] = Convert.ToDouble(xlSheet.Cells[X, 5].Value.ToString().Trim());
                                        }
                                        if (xlSheet.Cells[X, 6].Text.ToString() != "")
                                        {
                                            try
                                            {
                                                row["Date_Added"] = Convert.ToDateTime(xlSheet.Cells[X, 6].Text);
                                            }
                                            catch (FormatException)
                                            {
                                                row["Date_Added"] = DateTime.UtcNow;
                                            }
                                        }
                                        row["Platform"] = platform;
                                        row["ODM"] = odm;
                                        row["Comment"] = xlSheet.Cells[X, 9 + passthru].Text.ToString() == "" ? null : xlSheet.Cells[X, 9 + passthru].Text.ToString().Trim();
                                        row["filepath"] = filepath;
                                        dtBUSA.Rows.Add(row);
                                    }
                                }
                                else
                                {
                                    rownum++;
                                    if (rownum >= 5)
                                    {
                                        break;
                                    }
                                }
                            }

                            //04: Platform Data                 
                            if (strPlatform != "")
                            {
                                strPlatform = strPlatform.Replace("W1", " W1").Replace("1", " 1");  //e.g. 12, 14, 15 especially for WaferR (AI platform)
                                string[] array = strPlatform.ToString().Trim().Split(' ');
                                if (array[0].Length <= 4)  //for B52, Ionian13, warpath  || tracking.Contains("#")
                                {
                                    CPCT.Rows[i]["Platform"] = array[0];
                                }
                                else if (strPlatform.ToString().Contains("6U"))
                                {
                                    if (array[0].Contains("6U"))
                                    {
                                        CPCT.Rows[i]["Platform"] = array[0];
                                    }
                                    else if ((array[0] + array[1]).Contains("6U"))
                                    {
                                        CPCT.Rows[i]["Platform"] = array[0] + array[1];
                                    }
                                }
                                else if (tracking.Contains("6U"))
                                {
                                    CPCT.Rows[i]["Platform"] = tracking;
                                }
                                else
                                {
                                    CPCT.Rows[i]["Platform"] = Regex.Replace(array[0], @"[\d-]", string.Empty).Replace(".", string.Empty).Replace("_", string.Empty);
                                }
                            }
                            else
                            {
                                CPCT.Rows[i]["Error_flag"] = 1;
                                CPCT.Rows[i]["ErrorMsg"] = "Cannot find Platform info in BUSA sheet.";
                                goto End;
                            }

                            //05: ODM Data (*replace IEC with Inventec)
                            if (strODM != "")
                            {
                                CPCT.Rows[i]["ODM"] = strODM.ToString().Trim();
                            }
                            else
                            {
                                if (ep.Workbook.Worksheets["Summary"] != null)
                                {
                                    for (int X = 1; X <= xlSheet.Dimension.End.Column; X++)
                                    {
                                        string cell_1 = xlSheet.Cells[1, X].Value == null ? "" : xlSheet.Cells[1, X].Value.ToString().Trim();
                                        string cell_2 = xlSheet.Cells[2, X].Value == null ? "" : xlSheet.Cells[2, X].Value.ToString().Replace("IEC", "Inventec").Replace("CQ", "").Trim();  //.Replace("IMX", "MX")

                                        if (cell_1.ToLower().Contains("name of odm"))
                                        {
                                            if (cell_2 != "")
                                            {
                                                CPCT.Rows[i]["ODM"] = cell_2;
                                                break;
                                            }
                                            else
                                            {
                                                CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in Summary sheet.";
                                                goto End;
                                            }
                                        }
                                    }
                                }
                                else if (xlSheet == null && CPCT.Rows[i]["MultipleODM"].ToString() == "1")
                                {
                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in Summary sheet.";
                                    goto End;
                                }
                                else
                                {
                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in BUSA sheet.";
                                    goto End;
                                }
                            }

                            //ep.Save();
                            //ep.Dispose();
                        }
                        else
                        {
                            CPCT.Rows[i]["Error_flag"] = 1;
                            CPCT.Rows[i]["ErrorMsg"] = "Cannot find BUSA sheet.";
                            continue;
                        }
                    }

                End:
                    fs.Close();
                    fs.Dispose();
                }
            }
            return dtBUSA;
        }

        static DataTable BUSA_XXX(DataTable CPCT, string ver)
        {
            DataTable dtBUSA = new DataTable();
            dtBUSA.Columns.Add("BUSA");
            dtBUSA.Columns.Add("Description");
            dtBUSA.Columns.Add("Total_Cost");
            dtBUSA.Columns.Add("ODM_Managed_Cost");
            dtBUSA.Columns.Add("VA");
            dtBUSA.Columns.Add("Date_Added");
            dtBUSA.Columns.Add("Platform");
            dtBUSA.Columns.Add("ODM");
            dtBUSA.Columns.Add("Comment");
            dtBUSA.Columns.Add("Flag");
            dtBUSA.Columns.Add("UploadTime");
            dtBUSA.Columns.Add("filepath");
            dtBUSA.Columns.Add("version");
            dtBUSA.Columns["ODM_Managed_Cost"].DefaultValue = 0;
            dtBUSA.Columns["VA"].DefaultValue = 0;
            dtBUSA.Columns["Flag"].DefaultValue = "NB";
            dtBUSA.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtBUSA.Columns["version"].DefaultValue = ver;

            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                string filename = CPCT.Rows[i]["filename"].ToString().Trim();
                string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                string tracking = CPCT.Rows[i]["trackingname"].ToString().Trim();

                using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet xlSheet = ep.Workbook.Worksheets["BU SA"];
                        if (xlSheet != null)
                        {
                            //01: Check headings (*check "PassThru" for Crux)
                            int passthru = 0;
                            string[] colname_1 = new string[] { "Level 3", "Description", "Base Unit Cost", "ODM Managed Cost", "VA", "Date", "Platform", "ODM" };
                            string[] colname_2 = new string[] { "Level 3", "Description", "Base Unit Cost", "ODM Managed Cost", "VA", "Date", "Platform", "Pass", "ODM" };

                            for (int X = 0; X < colname_1.Length; X++)
                            {
                                if (xlSheet.Cells[1, 1 + X].Value != null && !xlSheet.Cells[1, 1 + X].Value.ToString().Contains(colname_1[X]))
                                {
                                    for (int Y = 0; Y < colname_2.Length; Y++)
                                    {
                                        passthru = 1;
                                        if (xlSheet.Cells[1, 1 + Y].Value != null && !xlSheet.Cells[1, 1 + Y].Value.ToString().Contains(colname_2[Y]))
                                        {
                                            CPCT.Rows[i]["Error_flag"] = 1;
                                            CPCT.Rows[i]["ErrorMsg"] = "The heading of BUSA sheet does not align with template.";
                                            goto End;
                                        }
                                    }
                                }
                            }

                            //02: Find Platform ODM (*replace IEC with Inventec)(*Multi ODM flag)(*allow 5 empty rows in between)
                            int rownum = 0;
                            string strPlatform = "", strODM = "";
                            for (int X = 2; X <= xlSheet.Dimension.End.Row; X++)
                            {
                                string platform = xlSheet.Cells[X, 7].Value == null ? "" : xlSheet.Cells[X, 7].Value.ToString().Trim();
                                string odm = xlSheet.Cells[X, 8 + passthru].Value == null ? "" : xlSheet.Cells[X, 8 + passthru].Value.ToString().Replace("IEC", "Inventec").Replace("CQ", "").Trim();  //.Replace("IMX", "MX")

                                if (platform != "" && odm != "")
                                {
                                    if (strPlatform == "" && strODM == "")
                                    {
                                        strPlatform = platform;
                                        strODM = odm;
                                    }
                                    else
                                    {
                                        if (platform.Length < strPlatform.Length)
                                        {
                                            strPlatform = platform;
                                        }
                                        if (odm != strODM)
                                        {
                                            strODM = "";
                                            CPCT.Rows[i]["MultipleODM"] = 1;
                                        }
                                    }

                                    if (xlSheet.Cells[X, 1].Value != null && xlSheet.Cells[X, 3].Value != null && Regex.IsMatch(xlSheet.Cells[X, 3].Value?.ToString(), @"^-?\d+"))  //+- digital number (more)
                                    {
                                        DataRow row = dtBUSA.NewRow();
                                        row["BUSA"] = xlSheet.Cells[X, 1].Value?.ToString().Trim();
                                        row["Description"] = xlSheet.Cells[X, 2].Value?.ToString().Trim();
                                        row["Total_Cost"] = Convert.ToDouble(xlSheet.Cells[X, 3].Value.ToString().Trim());

                                        if (xlSheet.Cells[X, 4].Value != null && Regex.IsMatch(xlSheet.Cells[X, 4].Value.ToString(), @"^-?\d+"))
                                        {
                                            row["ODM_Managed_Cost"] = Convert.ToDouble(xlSheet.Cells[X, 4].Value.ToString().Trim());
                                        }
                                        if (xlSheet.Cells[X, 5].Value != null && Regex.IsMatch(xlSheet.Cells[X, 5].Value.ToString(), @"^-?\d+"))
                                        {
                                            row["VA"] = Convert.ToDouble(xlSheet.Cells[X, 5].Value.ToString().Trim());
                                        }
                                        if (xlSheet.Cells[X, 6].Text.ToString() != "")
                                        {
                                            try
                                            {
                                                row["Date_Added"] = Convert.ToDateTime(xlSheet.Cells[X, 6].Text);
                                            }
                                            catch (FormatException)
                                            {
                                                row["Date_Added"] = DateTime.UtcNow;
                                            }
                                        }
                                        row["Platform"] = platform;
                                        row["ODM"] = odm;
                                        row["Comment"] = xlSheet.Cells[X, 9 + passthru].Value == null ? null : xlSheet.Cells[X, 9 + passthru].Value.ToString().Trim();
                                        row["filepath"] = filepath;
                                        dtBUSA.Rows.Add(row);
                                    }
                                }
                                else
                                {
                                    rownum++;
                                    if (rownum >= 5)
                                    {
                                        break;
                                    }
                                }
                            }

                            //04: Platform Data (*6U for andaman 6U)
                            if (strPlatform != "")
                            {
                                string[] array = strPlatform.ToString().Trim().Split(' ');
                                if (array[0].Length <= 4 || tracking.Contains("#"))  //for B52, Ionian13, warpath
                                {
                                    CPCT.Rows[i]["Platform"] = array[0];
                                }
                                else if (strPlatform.ToString().Contains("6U"))
                                {
                                    if (array[0].Contains("6U"))
                                    {
                                        CPCT.Rows[i]["Platform"] = array[0];
                                    }
                                    else if ((array[0] + array[1]).Contains("6U"))
                                    {
                                        CPCT.Rows[i]["Platform"] = array[0] + array[1];
                                    }
                                }
                                else if (tracking.Contains("6U"))
                                {
                                    CPCT.Rows[i]["Platform"] = tracking;
                                }
                                else
                                {
                                    CPCT.Rows[i]["Platform"] = Regex.Replace(array[0], @"[\d-]", string.Empty).Replace(".", string.Empty).Replace("_", string.Empty);
                                }
                            }
                            else
                            {
                                CPCT.Rows[i]["Error_flag"] = 1;
                                CPCT.Rows[i]["ErrorMsg"] = "Cannot find Platform info in BUSA sheet.";
                                goto End;
                            }

                            //05: ODM Data (*replace IEC with Inventec)
                            if (strODM != "")
                            {
                                CPCT.Rows[i]["ODM"] = strODM.ToString().Trim();
                            }
                            else
                            {
                                if (ep.Workbook.Worksheets["Summary"] != null)
                                {
                                    for (int X = 1; X <= xlSheet.Dimension.End.Column; X++)
                                    {
                                        string cell_1 = xlSheet.Cells[1, X].Value == null ? "" : xlSheet.Cells[1, X].Value.ToString().Trim();
                                        string cell_2 = xlSheet.Cells[2, X].Value == null ? "" : xlSheet.Cells[2, X].Value.ToString().Replace("IEC", "Inventec").Replace("CQ", "").Trim();  //.Replace("IMX", "MX")

                                        if (cell_1.ToLower().Contains("name of odm"))
                                        {
                                            if (cell_2 != "")
                                            {
                                                CPCT.Rows[i]["ODM"] = cell_2;
                                                break;
                                            }
                                            else
                                            {
                                                CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in Summary sheet.";
                                                goto End;
                                            }
                                        }
                                    }
                                }
                                else if (xlSheet == null && CPCT.Rows[i]["MultipleODM"].ToString() == "1")
                                {
                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in Summary sheet.";
                                    goto End;
                                }
                                else
                                {
                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot find ODM info in BUSA sheet.";
                                    goto End;
                                }
                            }

                            //ep.Save();
                            //ep.Dispose();
                        }
                        else
                        {
                            CPCT.Rows[i]["Error_flag"] = 1;
                            CPCT.Rows[i]["ErrorMsg"] = "Cannot find BUSA sheet.";
                            continue;
                        }
                    }

                End:
                    fs.Close();
                    fs.Dispose();
                }
            }
            return dtBUSA;
        }

        static DataTable UploadBUSA(SqlConnection conn, DataTable CPCT, DataTable dtBUSA, string ver)
        {
            //01: Upload BUSA
            SqlBulkCopy sbcBUSA = new SqlBulkCopy(conn);
            sbcBUSA.DestinationTableName = "[dbo].[Eevee_BUSA]";
            sbcBUSA.ColumnMappings.Add("BUSA", "BUSA");
            sbcBUSA.ColumnMappings.Add("Description", "Description");
            sbcBUSA.ColumnMappings.Add("Total_Cost", "Total_Cost");
            sbcBUSA.ColumnMappings.Add("ODM_Managed_Cost", "ODM_Managed_Cost");
            sbcBUSA.ColumnMappings.Add("VA", "VA");
            sbcBUSA.ColumnMappings.Add("Date_Added", "Date_Added");
            sbcBUSA.ColumnMappings.Add("Platform", "Platform");
            sbcBUSA.ColumnMappings.Add("ODM", "ODM");
            sbcBUSA.ColumnMappings.Add("Comment", "Comment");
            sbcBUSA.ColumnMappings.Add("Flag", "Flag");
            sbcBUSA.ColumnMappings.Add("UploadTime", "UploadTime");
            sbcBUSA.ColumnMappings.Add("filepath", "filepath");
            sbcBUSA.ColumnMappings.Add("version", "version");
            sbcBUSA.BulkCopyTimeout = 0;
            sbcBUSA.WriteToServer(dtBUSA);
            sbcBUSA.Close();

            dtBUSA.Clear();
            dtBUSA.Dispose();

            //02: Capture PL info, AV site, PMatrix SA Desc/Category
            string strUpdate =
                "UPDATE t1 SET [PL]=t2.[PL] FROM [dbo].[Eevee_BUSA] t1, (SELECT Distinct [PL], [SA_Number] FROM [dbo].[PMatrix_AV_SA] UNION SELECT Distinct [PL], [Comp_Number] FROM [dbo].[PMatrix_SA_Comp]) t2 WHERE t1.[BUSA]=t2.[SA_Number] and t2.[PL] is not null and t1.[version]='" + ver + "' " +
                "UPDATE t1 SET [AV_Site_Committed]=t2.[AV_Site] FROM [dbo].[Eevee_BUSA] t1, (SELECT Distinct [BUSA], [AV_Site] FROM [dbo].[BB8_AV_Site_Committed]) t2 WHERE t1.[BUSA]=t2.[BUSA] and t1.[version]='" + ver + "' " +
                "UPDATE t1 SET [PMatrix_Family]=LTRIM(t2.[Family]), [PMatrix_SA_Desc]=LTRIM(t2.[SA_Description]), [PMatrix_Category]=t2.[Category] FROM [dbo].[Eevee_BUSA] t1, (SELECT Distinct [Family], [SA_Number], [SA_Description], [Category] FROM [dbo].[PMatrix_AV_SA]) t2 WHERE t1.[BUSA]=t2.[SA_Number] and t1.[version]='" + ver + "' " +  //for IDS
                "UPDATE t1 SET [PMatrix_Family]=LTRIM(t2.[Family]), [PMatrix_SA_Desc]=LTRIM(t2.[Comp_Description]), [PMatrix_Category]=t2.[Category] FROM [dbo].[Eevee_BUSA] t1, (SELECT Distinct [Family], [Comp_Number], [Comp_Description], [Category] FROM [dbo].[PMatrix_SA_Comp]) t2 WHERE t1.[BUSA]=t2.[Comp_Number] and t1.[version]='" + ver + "'";  //for RCTO
            SqlCommand cmdUpdate = new SqlCommand(strUpdate, conn);
            cmdUpdate.CommandTimeout = 0;
            cmdUpdate.ExecuteNonQuery();

            //missing: PL for IDS BUSA
            //missing: AV site committe for RCTO BUSA
            //missing: non BUSA in CPCT
            //missing: MLCC adder

            string FCYear = "", FCQuarter = "";
            if (DateTime.UtcNow.Month >= 11)
            {
                FCYear = "FY" + DateTime.UtcNow.AddYears(1).Year.ToString().Substring(2, 2);
                FCQuarter = "Q1";
            }
            else if (DateTime.UtcNow.Month == 1)
            {
                FCYear = "FY" + DateTime.UtcNow.Year.ToString().Substring(2, 2);
                FCQuarter = "Q1";
            }
            else if (DateTime.UtcNow.Month <= 4)
            {
                FCYear = "FY" + DateTime.UtcNow.Year.ToString().Substring(2, 2);
                FCQuarter = "Q2";
            }
            else if (DateTime.UtcNow.Month <= 7)
            {
                FCYear = "FY" + DateTime.UtcNow.Year.ToString().Substring(2, 2);
                FCQuarter = "Q3";
            }
            else
            {
                FCYear = "FY" + DateTime.UtcNow.Year.ToString().Substring(2, 2);
                FCQuarter = "Q4";
            }

            //03: Refresh: MLCC, P10, Resilient adder, RCTO adder
            //DataTable dtMLCC = new DataTable();
            //string sqlMLCC = "SELECT * FROM [dbo].[BB8_Adder_MLCC] WHERE [Fiscal_Year]='" + FCYear + "' and [Fiscal_Quarter]='" + FCQuarter + "'";
            //SqlDataAdapter sdaMLCC = new SqlDataAdapter(sqlMLCC, conn);
            //sdaMLCC.Fill(dtMLCC);

            //DataTable dtP10 = new DataTable();
            //string sqlP10 = "SELECT * FROM [dbo].[BB8_Adder_P10] WHERE [Fiscal_Year]='" + FCYear + "'";
            //SqlDataAdapter sdaP10 = new SqlDataAdapter(sqlP10, conn);
            //sdaP10.Fill(dtP10);

            string strReset = "UPDATE [dbo].[Eevee_BUSA] SET [Adder_MLCC]=0, [Adder_P10]=0, [Adder_Resilient]=0, [Adder_RCTO]=0, [Adder_PassThru]=0 WHERE [version]='" + ver + "'";
            SqlCommand cmdReset = new SqlCommand(strReset, conn);
            cmdReset.ExecuteNonQuery();

            string strAdder =
                "UPDATE t1 SET [Adder_MLCC]=t2.[MLCC] FROM [dbo].[Eevee_BUSA] t1, [dbo].[BB8_Adder_MLCC] t2 " +
                "WHERE t1.[version]='" + ver + "' and [Fiscal_Year]='" + FCYear + "' and [Fiscal_Quarter]='" + FCQuarter + "' and [PMatrix_Category] LIKE '%Base Unit%' and [PMatrix_SA_Desc] NOT LIKE '%Cost adder%' " +  //([PMatrix_Category] LIKE '%Base Unit%' OR [Description] LIKE '%BU%')
                "" +
                "UPDATE t1 SET [Adder_P10]=t2.[P10_Adder] FROM [dbo].[Eevee_BUSA] t1, [dbo].[BB8_Adder_P10] t2 " +
                "WHERE t1.[PL]=t2.[PL] and t1.[version]='" + ver + "' and t2.[Fiscal_Year]='" + FCYear + "' and [PMatrix_Category] LIKE '%Base Unit%' and [PMatrix_SA_Desc] NOT LIKE '%RCT%' and [PMatrix_SA_Desc] NOT LIKE '%Cost adder%' " +
                "" +
                "UPDATE tb1 SET [Adder_Resilient]=tb2.[delta] FROM [dbo].[Eevee_BUSA] tb1, (SELECT t1.[Resilient], t2.*, t3.[Total_Cost] as [min_cost], t3.[Total_Cost]-t2.[Total_Cost] as [delta] FROM " +
                "(SELECT [BUSA], COUNT([Total_Cost]) as [Resilient] FROM (SELECT Distinct [BUSA], [Total_Cost] FROM [dbo].[Eevee_BUSA] WHERE [version]='" + ver + "') A GROUP BY [BUSA]) t1 " +
                "LEFT JOIN [dbo].[Eevee_BUSA] as t2 on t2.[BUSA]=t1.[BUSA] " +
                "LEFT JOIN (SELECT *, ROW_NUMBER() OVER (PARTITION BY [BUSA] ORDER BY [Total_Cost]) as [srank] FROM [dbo].[Eevee_BUSA] WHERE [version]='" + ver + "') as t3 on t3.[BUSA]=t1.[BUSA] " +
                "WHERE t2.[version]='" + ver + "' and t2.[PMatrix_Category] LIKE '%Base Unit%' and (t2.[PMatrix_SA_Desc] NOT LIKE '%RCT%' OR t2.[Description] NOT LIKE '%RCT%') and [Resilient]>1 and t3.[srank]=1) tb2 WHERE tb1.[int_BUSA_ID]=tb2.[int_BUSA_ID] " +
                "" +
                "UPDATE t1 SET [Adder_RCTO]=t2.[RCTO_Adder] FROM [dbo].[Eevee_BUSA] t1, [dbo].[BB8_Adder_RCTO] t2 " +
                "WHERE t1.[version]='" + ver + "' and t1.[AV_Site_Committed]=t2.[AV_Committed] and t1.[ODM] LIKE ''+t2.[ODM]+'%' " +
                "" +
                "UPDATE t1 SET [Adder_PassThru]=3.18 FROM [dbo].[Eevee_BUSA] t1 " +   //Adder Pass Thru
                "WHERE [version]='" + ver + "' and [ODM] in ('Huaqin', 'FXN LH') and [PMatrix_Category] LIKE 'Base Unit%' and [PMatrix_SA_Desc] NOT LIKE '%RCT%' and [PMatrix_SA_Desc] NOT LIKE '%Cost adder%'";
            SqlCommand cmdAdder = new SqlCommand(strAdder, conn);
            cmdAdder.CommandTimeout = 0;
            cmdAdder.ExecuteNonQuery();

            string strSUM = "UPDATE [dbo].[Eevee_BUSA] SET [Total_Adder]=[Adder_MLCC]+[Adder_P10]+[Adder_Resilient]+[Adder_RCTO]+[Adder_PassThru] WHERE [version]='" + ver + "'";
            SqlCommand cmdSUM = new SqlCommand(strSUM, conn);
            cmdSUM.ExecuteNonQuery();

            //02: Remove EOL Platform
            DataTable dtEOL = new DataTable();
            string strEOL = "SELECT Distinct [filename] FROM [dbo].[Eevee_EOL] WHERE [PL]='NB' and [Valid]='Y'";
            SqlDataAdapter sdaEOL = new SqlDataAdapter(strEOL, conn);
            sdaEOL.Fill(dtEOL);

        CPC_Remove:
            for (int X = 0; X < CPCT.Rows.Count; X++)
            {
                for (int Y = 0; Y < dtEOL.Rows.Count; Y++)
                {
                    if (CPCT.Rows[X]["filepath"].ToString().Contains(dtEOL.Rows[Y]["filename"].ToString().Trim()))
                    {
                        CPCT.Rows[X].Delete();
                        goto CPC_Remove;
                    }
                }
            }

            //03: Upload CPCT History (By "version")
            SqlBulkCopy SBC = new SqlBulkCopy(conn);
            SBC.DestinationTableName = "[dbo].[Eevee_File_History]";
            SBC.ColumnMappings.Add("filename", "filename");
            SBC.ColumnMappings.Add("trackingname", "trackingname");
            SBC.ColumnMappings.Add("filepath", "filepath");
            SBC.ColumnMappings.Add("UploadTime", "UploadTime");
            SBC.ColumnMappings.Add("Platform", "Platform");
            SBC.ColumnMappings.Add("ODM", "ODM");
            SBC.ColumnMappings.Add("Error_flag", "Error_flag");
            SBC.ColumnMappings.Add("ErrorMsg", "ErrorMsg");
            SBC.ColumnMappings.Add("MultipleODM", "MultipleODM");
            SBC.ColumnMappings.Add("version", "version");
            SBC.ColumnMappings.Add("CurrentMonth", "CurrentMonth");
            SBC.BulkCopyTimeout = 0;
            SBC.WriteToServer(CPCT);
            SBC.Close();

            string strBUSAnAdder =
                "INSERT INTO [dbo].[Eevee_Upload_Report]([PN], [Sitecode], [Vendorcode], [Cost], [EffDateFrom], [EffDateTo], [BU], [Type], [UploadTime], [filepath], [version]) " +
                "SELECT [BUSA], [Site_Code], [Vendor_Code], [Total_Cost], GETUTCDATE(), '9999-12-31', 'NB', 'BU' as [Type], GETUTCDATE(), A.[filepath], A.[version] FROM [dbo].[Eevee_BUSA] A " +
                "LEFT JOIN [dbo].[BB8_ODM_Lookup] as B on B.[ODM]=A.[ODM] " +
                "WHERE [BUSA] is not NULL and B.[Note]='IDS' and B.[BU]='NB' and [PMatrix_Category]='Base Unit' and A.[version]='" + ver + "' " +
                "UNION " +
                "SELECT [BUSA], [Site_Code], [Vendor_Code], [Total_Adder], GETUTCDATE(), '9999-12-31', 'NB', 'Adder' as [Type], GETUTCDATE(), A.[filepath], A.[version] FROM [dbo].[Eevee_BUSA] A " +
                "LEFT JOIN [dbo].[BB8_ODM_Lookup] as B on B.[ODM]=A.[ODM] " +
                "WHERE [BUSA] is not NULL and [PMatrix_Category] LIKE '%Base Unit%' and [PMatrix_SA_Desc] NOT LIKE '%Cost adder%' and B.[Note]='IDS' and B.[BU]='NB' and A.[version]='" + ver + "' " +
                "UNION " +
                "SELECT [BUSA], [Site_Code], [Vendor_Code], [Total_Cost], GETUTCDATE(), '9999-12-31', 'NB', 'Option' as [Type], GETUTCDATE(), A.[filepath], A.[version] FROM [dbo].[Eevee_BUSA] A " +
                "LEFT JOIN [dbo].[BB8_ODM_Lookup] as B on B.[ODM]=A.[ODM] " +
                "WHERE [BUSA] is not NULL and B.[Note]='IDS' and B.[BU]='NB' and [PMatrix_Category]<>'Base Unit' and A.[version]='" + ver + "' " +
                "order by [filepath], [BUSA], [Type] desc";
            SqlCommand cmdBUSAnAdder = new SqlCommand(strBUSAnAdder, conn);
            cmdBUSAnAdder.ExecuteNonQuery();

            //06: Find PMatrix Hierarchy
            DataTable dtPMatrix = new DataTable();
            string strPMatrix = "SELECT Distinct [Family] FROM [dbo].[PMatrix_Family]";
            SqlDataAdapter sdaPMatrix = new SqlDataAdapter(strPMatrix, conn);
            sdaPMatrix.Fill(dtPMatrix);

            DataView dv = new DataView(dtPMatrix);

            DataTable dtHierarchy = new DataTable();
            dtHierarchy.Columns.Add("Platform");
            dtHierarchy.Columns.Add("filepath");

            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                string platform = CPCT.Rows[i]["Platform"].ToString().Trim();

                if (CPCT.Rows[i]["Error_flag"].ToString() == "")
                {
                    dv.RowFilter = string.Empty;
                    dv.RowFilter = "Family LIKE '" + platform + "%'";

                    for (int j = 0; j < dv.ToTable().Rows.Count; j++)
                    {
                        if (platform.Length < 4)  //Pan, Roc
                        {
                            int checkpmatrix = 0;
                            string[] array = dv.ToTable().Rows[j]["Family"].ToString().ToString().Trim().Split(' ');
                            string family = array[0].Replace(platform, "").Replace("R6U", "").Replace("RPL6U", "").Replace("RT6U", "").Trim();   //"R6U", "RPL6U", "RT6U"

                            foreach (var letter in family.ToCharArray())
                            {
                                if (char.IsLower(letter) || char.IsUpper(letter))
                                {
                                    checkpmatrix = 1;
                                    break;
                                }
                            }

                            if (checkpmatrix == 0)
                            {
                                DataRow row = dtHierarchy.NewRow();
                                row["Platform"] = dv.ToTable().Rows[j]["Family"].ToString().Trim();
                                row["filepath"] = filepath;
                                dtHierarchy.Rows.Add(row);
                            }
                        }
                        else
                        {

                            DataRow row = dtHierarchy.NewRow();
                            row["Platform"] = dv.ToTable().Rows[j]["Family"].ToString().Trim();
                            row["filepath"] = filepath;
                            dtHierarchy.Rows.Add(row);
                        }
                    }
                }
            }
            return dtHierarchy;
        }

        public static Tuple<DataTable, DataTable, DataTable, DataTable, DataTable> Option(DataTable CPCT, DataTable dtHierarchy, string[] Extend6U, string ver)
        {
            DataTable dtOption = new DataTable();
            dtOption.Columns.Add("ODM");
            dtOption.Columns.Add("Platform");
            dtOption.Columns.Add("Size");
            dtOption.Columns.Add("Extension"); //**for R6U, RPL6U 
            dtOption.Columns.Add("Config_Option");
            dtOption.Columns.Add("Config_Num");
            dtOption.Columns.Add("Config_Rule");
            dtOption.Columns.Add("Rule_1");
            dtOption.Columns.Add("Rule_2");
            dtOption.Columns.Add("Rule_3");
            dtOption.Columns.Add("AV_Category");
            dtOption.Columns.Add("Cost");
            dtOption.Columns.Add("rowId");
            dtOption.Columns.Add("filesheet");
            dtOption.Columns.Add("filepath");
            dtOption.Columns.Add("UploadTime");
            dtOption.Columns.Add("version");
            dtOption.Columns["Cost"].DefaultValue = 0;
            dtOption.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtOption.Columns["version"].DefaultValue = ver;

            DataTable dtRCTO = new DataTable();
            dtRCTO.Columns.Add("Platform");
            dtRCTO.Columns.Add("ODM");
            dtRCTO.Columns.Add("RCTO_Site");
            dtRCTO.Columns.Add("filepath");
            dtRCTO.Columns.Add("filesheet");
            dtRCTO.Columns.Add("UploadTime");
            dtRCTO.Columns.Add("version");
            dtRCTO.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtRCTO.Columns["version"].DefaultValue = ver;

            DataTable dtRCTO_PN = new DataTable();
            dtRCTO_PN.Columns.Add("ODM");
            dtRCTO_PN.Columns.Add("Platform");
            dtRCTO_PN.Columns.Add("Config_Rule");
            dtRCTO_PN.Columns.Add("rowId");
            dtRCTO_PN.Columns.Add("RCTO_Site");
            dtRCTO_PN.Columns.Add("PMatrix");
            dtRCTO_PN.Columns.Add("filepath");
            dtRCTO_PN.Columns.Add("filesheet");
            dtRCTO_PN.Columns.Add("UploadTime");
            dtRCTO_PN.Columns.Add("version");
            dtRCTO_PN.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtRCTO_PN.Columns["version"].DefaultValue = ver;

            DataTable dtConfig_PN = new DataTable();
            dtConfig_PN.Columns.Add("ODM");
            dtConfig_PN.Columns.Add("Platform");
            dtConfig_PN.Columns.Add("Size");
            dtConfig_PN.Columns.Add("Extension");
            dtConfig_PN.Columns.Add("rowId");
            dtConfig_PN.Columns.Add("Config_Rule");
            dtConfig_PN.Columns.Add("AV_Category");
            dtConfig_PN.Columns.Add("SA_PN");
            dtConfig_PN.Columns.Add("Cost");
            dtConfig_PN.Columns.Add("filesheet");
            dtConfig_PN.Columns.Add("filepath");
            dtConfig_PN.Columns.Add("UploadTime");
            dtConfig_PN.Columns.Add("version");
            dtConfig_PN.Columns["Cost"].DefaultValue = 0;
            dtConfig_PN.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtConfig_PN.Columns["version"].DefaultValue = ver;

            DataTable dtExclude = new DataTable();
            dtExclude.Columns.Add("Platform");
            dtExclude.Columns.Add("Exclude");
            dtExclude.Columns.Add("filepath");
            dtExclude.Columns.Add("filesheet");

            DataView dvHierarchy = new DataView(dtHierarchy);

            int numSheet = 0;
            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                string filename = CPCT.Rows[i]["filename"].ToString().Trim();
                string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                string platform = CPCT.Rows[i]["Platform"].ToString().Trim();

                if (CPCT.Rows[i]["Error_flag"].ToString() == "")
                {
                    try
                    {
                        using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            using (ExcelPackage ep = new ExcelPackage(fs))
                            {
                                foreach (ExcelWorksheet xlSheet in ep.Workbook.Worksheets)
                                {
                                    string sheet_name = xlSheet.Name.ToString().Trim();
                                    if (sheet_name.StartsWith("Summary"))
                                    {
                                        numSheet++;

                                        //01: Find Keyword
                                        string tempSize = "", strODM = "";
                                        int initCol = 0, cateCol = 0, cateRow = 0, RCTOforPNCol = 0, ConfigforPNCol = 0;
                                        List<List<string>> SizeColList = new List<List<string>>();  //Size, column number

                                        for (int X = 1; X <= xlSheet.Dimension.End.Row; X++)
                                        {
                                            for (int Y = 1; Y <= xlSheet.Dimension.End.Column; Y++)
                                            {
                                                if (xlSheet.Cells[X, Y].Value != null)
                                                {
                                                    //001: Different Size + Col Number
                                                    if (strODM == "" && xlSheet.Cells[X, Y].Value.ToString().Contains("Name of ODM"))
                                                    {
                                                        strODM = xlSheet.Cells[2, Y].Value.ToString().Replace("IEC", "Inventec").Replace("CQ", "").Trim();  //.Replace("IMX", "MX")
                                                    }
                                                    else if (initCol == 0 && xlSheet.Cells[X, Y].Value.ToString().Contains("Description"))
                                                    {
                                                        initCol = Y;

                                                        for (int Z = initCol + 1; Z <= xlSheet.Dimension.End.Column; Z++)
                                                        {
                                                            if (xlSheet.Cells[X, Z].Value != null && xlSheet.Cells[X, Z].Value.ToString().Contains("Original"))
                                                            {
                                                                if (xlSheet.Cells[X - 2, Z].Value != null) //Size headers
                                                                {
                                                                    tempSize = xlSheet.Cells[X - 2, Z].Value.ToString().Replace('"', ' ').Replace("'", "").Replace(platform, "").Trim();

                                                                    List<string> size = new List<string>();
                                                                    size.Add(tempSize);
                                                                    size.Add((Z + 2).ToString());
                                                                    SizeColList.Add(size);
                                                                }
                                                                else
                                                                {
                                                                    if (tempSize == "")  //check the first architect
                                                                    {
                                                                        CPCT.Rows[i]["Error_flag"] = 1;
                                                                        CPCT.Rows[i]["ErrorMsg"] = "Architecture size cannot be empty on row " + (X - 2) + ".";
                                                                        goto End;
                                                                    }
                                                                    else
                                                                    {
                                                                        List<string> size = new List<string>();  //for those merged cells 
                                                                        size.Add(tempSize);
                                                                        size.Add((Z + 2).ToString());
                                                                        SizeColList.Add(size);
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        X += 10;
                                                        Y = xlSheet.Dimension.End.Column;
                                                    }
                                                    else if (X == 1 && xlSheet.Cells[X + 1, Y].Value != null && xlSheet.Cells[X, Y].Value.ToString().ToLower().Contains("include"))
                                                    {
                                                        if (xlSheet.Cells[X + 1, Y].Value.ToString().Contains(","))
                                                        {
                                                            string[] array = xlSheet.Cells[X + 1, Y].Value.ToString().Split(',');
                                                            for (int Z = 0; Z < array.Length; Z++)
                                                            {
                                                                DataRow row = dtExclude.NewRow();
                                                                row["Platform"] = platform;
                                                                row["Exclude"] = array[Z].Trim();
                                                                row["filepath"] = filepath;
                                                                row["filesheet"] = sheet_name;
                                                                dtExclude.Rows.Add(row);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            DataRow row = dtExclude.NewRow();
                                                            row["Platform"] = platform;
                                                            row["Exclude"] = xlSheet.Cells[X + 1, Y].Value.ToString().Trim();
                                                            row["filepath"] = filepath;
                                                            row["filesheet"] = sheet_name;
                                                            dtExclude.Rows.Add(row);
                                                        }
                                                    }
                                                    else if (X == 1 && xlSheet.Cells[X + 1, Y].Value != null && xlSheet.Cells[X, Y].Value.ToString().ToUpper().Contains("RCTO"))
                                                    {
                                                        if (xlSheet.Cells[X + 1, Y].Value.ToString().Contains(","))
                                                        {
                                                            string[] array = xlSheet.Cells[X + 1, Y].Value.ToString().Split(',');
                                                            for (int Z = 0; Z < array.Length; Z++)
                                                            {
                                                                DataRow row = dtRCTO.NewRow();
                                                                row["Platform"] = platform;
                                                                row["ODM"] = strODM;
                                                                row["filepath"] = filepath;
                                                                row["filesheet"] = sheet_name;
                                                                row["RCTO_Site"] = array[Z].Trim();
                                                                dtRCTO.Rows.Add(row);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            DataRow row = dtRCTO.NewRow();
                                                            row["Platform"] = platform;
                                                            row["ODM"] = strODM;
                                                            row["filepath"] = filepath;
                                                            row["filesheet"] = sheet_name;
                                                            row["RCTO_Site"] = xlSheet.Cells[X + 1, Y].Value.ToString().Trim();
                                                            dtRCTO.Rows.Add(row);
                                                        }
                                                        CPCT.Rows[i]["RCTO_flag"] = 1;
                                                    }
                                                    else if (cateCol == 0 && cateRow == 0 && xlSheet.Cells[X, Y].Value.ToString().Contains("Category"))
                                                    {
                                                        cateCol = Y;
                                                        cateRow = X;
                                                    }
                                                    else if (xlSheet.Cells[X, Y].Value.ToString().Contains("@hp.com"))
                                                    {
                                                        CPCT.Rows[i]["Owner"] = xlSheet.Cells[X, Y].Value.ToString().ToLower().Trim();
                                                    }
                                                    else if (xlSheet.Cells[X, Y].Value.ToString().ToUpper().Contains("RCTO") && xlSheet.Cells[X, Y].Value.ToString().ToUpper().Contains("PN"))
                                                    {
                                                        RCTOforPNCol = Y;
                                                        CPCT.Rows[i]["RCTO_PN_flag"] = 1;                                                       
                                                    }
                                                    else if (xlSheet.Cells[X, Y].Value.ToString().ToLower().Contains("configurable for pn"))
                                                    {
                                                        ConfigforPNCol = Y;
                                                        goto Check_Keyword;
                                                    }
                                                }
                                            }
                                        }

                                    Check_Keyword:
                                        if (initCol == 0)
                                        {
                                            CPCT.Rows[i]["Error_flag"] = 1;
                                            CPCT.Rows[i]["ErrorMsg"] = "Cannot find the string 'Description' in " + sheet_name + " sheet";
                                            goto End;
                                        }
                                        if (cateCol == 0 && cateRow == 0)
                                        {
                                            CPCT.Rows[i]["Error_flag"] = 1;
                                            CPCT.Rows[i]["ErrorMsg"] = "Cannot find the 'Category' column in " + sheet_name + " sheet.";
                                            goto End;
                                        }

                                        dvHierarchy.RowFilter = string.Empty;
                                        dvHierarchy.RowFilter = "filepath LIKE '" + filepath + "%'";


                                        ////ODM!!!!
                                        //if (strODM.Contains(","))
                                        //{
                                        //    string[] ODM = strODM.ToString().Split(',');

                                        //    string test = "";
                                        //}
                                        //else
                                        //{
                                        //    string[] ODM = new string[1];
                                        //    ODM[0] = strODM;

                                        //    string test = "";
                                        //}
                                    
  
                                        //02: Option
                                        int rowId = 1;
                                        for (int j = cateRow + 1; j <= xlSheet.Dimension.End.Row; j++)
                                        {
                                            if (xlSheet.Cells[j, initCol].Value != null && xlSheet.Cells[j, cateCol - 1].Value != null && xlSheet.Cells[j, cateCol].Value != null)
                                            {
                                                //001: Skip Category "Base unit", "BU" 
                                                if (!xlSheet.Cells[j, cateCol].Value.ToString().ToLower().Contains("base") && !xlSheet.Cells[j, cateCol].Value.ToString().Contains("BU"))
                                                {
                                                    //002: RCTO for PN
                                                    if (RCTOforPNCol != 0 && xlSheet.Cells[j, RCTOforPNCol].Value != null)  //&& CPCT.Rows[i]["RCTO_PN_flag"].ToString() == "1"
                                                    {
                                                        if (xlSheet.Cells[j, RCTOforPNCol].Value.ToString().Contains(","))
                                                        {
                                                            string[] RCTO_for_PN = xlSheet.Cells[j, RCTOforPNCol].Value.ToString().Split(',');
                                                            for (int X = 0; X < RCTO_for_PN.Length; X++)
                                                            {
                                                                string[] RCTO_PN = RCTO_for_PN[X].Trim().Split(new string[2] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                                                                if (RCTO_PN.Length == 2)
                                                                {
                                                                    DataRow row = dtRCTO_PN.NewRow();
                                                                    row["Platform"] = platform;
                                                                    row["ODM"] = strODM;
                                                                    row["filepath"] = filepath;
                                                                    row["filesheet"] = sheet_name;
                                                                    row["rowId"] = rowId;
                                                                    row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                    row["PMatrix"] = RCTO_PN[0].ToString().Trim();
                                                                    row["RCTO_Site"] = RCTO_PN[1].ToString().Trim();
                                                                    dtRCTO_PN.Rows.Add(row);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string[] RCTO_PN = xlSheet.Cells[j, cateCol + 2].Value.ToString().Split(new string[2] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                                                            if (RCTO_PN.Length == 2)
                                                            {
                                                                DataRow row = dtRCTO_PN.NewRow();
                                                                row["Platform"] = platform;
                                                                row["ODM"] = strODM;
                                                                row["filepath"] = filepath;
                                                                row["filesheet"] = sheet_name;
                                                                row["rowId"] = rowId;
                                                                row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                row["PMatrix"] = RCTO_PN[0].ToString().Trim();
                                                                row["RCTO_Site"] = RCTO_PN[1].ToString().Trim();
                                                                dtRCTO_PN.Rows.Add(row);
                                                            }
                                                        }
                                                    }

                                                    //003: Configurable for PN
                                                    if (ConfigforPNCol !=0 && xlSheet.Cells[j, ConfigforPNCol].Value != null)
                                                    {
                                                        if (xlSheet.Cells[j, ConfigforPNCol].Value.ToString().Contains(","))
                                                        {
                                                            string[] Config_for_PN = xlSheet.Cells[j, ConfigforPNCol].Value.ToString().Split(',');
                                                            for (int X = 0; X < Config_for_PN.Length; X++)
                                                            {
                                                                //003.1 Loop by different size
                                                                for (int Z = 0; Z < dvHierarchy.ToTable().Rows.Count; Z++)
                                                                {
                                                                    string[] arrHierarchy = dvHierarchy.ToTable().Rows[Z]["Platform"].ToString().Trim().Split(' ');
                                                                    string strHierarchy = arrHierarchy[0].ToString().ToUpper().Replace(platform.ToUpper(), string.Empty);

                                                                    DataRow row = dtConfig_PN.NewRow();
                                                                    row["ODM"] = strODM;
                                                                    row["Platform"] = platform;
                                                                    row["Size"] = strHierarchy;
                                                                    row["Extension"] = "";  //6U, R6U, RPL6U
                                                                    row["filesheet"] = sheet_name;
                                                                    row["filepath"] = filepath;
                                                                    row["rowId"] = rowId;
                                                                    row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                    row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                    row["SA_PN"] = Config_for_PN[X].Trim();

                                                                    //Cost by Size
                                                                    foreach (List<string> innerlist in SizeColList)
                                                                    {
                                                                        string size = innerlist[0].ToString().Trim();
                                                                        int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                        string hierarchy = strHierarchy;

                                                                        if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                        {
                                                                            hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                            if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                            {
                                                                                if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                {
                                                                                    row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                        else //for 13, 14, 15
                                                                        {
                                                                            if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                            {
                                                                                if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                {
                                                                                    row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    dtConfig_PN.Rows.Add(row);

                                                                    //Extension for 6U, R6U, RPL6U, RT6U
                                                                    if (platform.ToUpper().Contains("6U"))
                                                                    {
                                                                        for (int Y = 0; Y < Extend6U.Length; Y++)
                                                                        {
                                                                            string[] arr6U = platform.ToUpper().Split(new string[] { "6U" }, StringSplitOptions.RemoveEmptyEntries);

                                                                            DataRow row6U = dtConfig_PN.NewRow();
                                                                            row6U["ODM"] = strODM;
                                                                            row6U["Platform"] = arr6U[0].Trim();
                                                                            row6U["Size"] = strHierarchy;
                                                                            row6U["Extension"] = Extend6U[Y];  //6U, R6U, RPL6U, RT6U
                                                                            row6U["filesheet"] = sheet_name;
                                                                            row6U["filepath"] = filepath;
                                                                            row6U["rowId"] = rowId;
                                                                            row6U["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                            row6U["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                            row6U["SA_PN"] = Config_for_PN[X].Trim();

                                                                            //Cost by Size
                                                                            foreach (List<string> innerlist in SizeColList)
                                                                            {
                                                                                string size = innerlist[0].ToString().Trim();
                                                                                int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                                string hierarchy = row6U["Size"].ToString();

                                                                                if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                                {
                                                                                    hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                                    if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                                    {
                                                                                        if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                        {
                                                                                            row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                            break;
                                                                                        }
                                                                                    }
                                                                                }
                                                                                else //for 13, 14, 15
                                                                                {
                                                                                    if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                                    {
                                                                                        if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                        {
                                                                                            row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                            break;
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            dtConfig_PN.Rows.Add(row6U);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //003.1 Loop by different size
                                                            for (int Z = 0; Z < dvHierarchy.ToTable().Rows.Count; Z++)
                                                            {
                                                                string[] arrHierarchy = dvHierarchy.ToTable().Rows[Z]["Platform"].ToString().Trim().Split(' ');
                                                                string strHierarchy = arrHierarchy[0].ToString().ToUpper().Replace(platform.ToUpper(), string.Empty);

                                                                DataRow row = dtConfig_PN.NewRow();
                                                                row["ODM"] = strODM;
                                                                row["Platform"] = platform;
                                                                row["Size"] = strHierarchy;
                                                                row["Extension"] = "";  //6U, R6U, RPL6U
                                                                row["filesheet"] = sheet_name;
                                                                row["filepath"] = filepath;
                                                                row["rowId"] = rowId;
                                                                row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                row["SA_PN"] = xlSheet.Cells[j, ConfigforPNCol].Value.ToString().Trim();

                                                                //Cost by Size
                                                                foreach (List<string> innerlist in SizeColList)
                                                                {
                                                                    string size = innerlist[0].ToString().Trim();
                                                                    int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                    string hierarchy = strHierarchy;

                                                                    if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                    {
                                                                        hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                        if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                        {
                                                                            if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                            {
                                                                                row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                    else //for 13, 14, 15
                                                                    {
                                                                        if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                        {
                                                                            if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                            {
                                                                                row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                dtConfig_PN.Rows.Add(row);

                                                                //Extension for 6U, R6U, RPL6U, RT6U
                                                                if (platform.ToUpper().Contains("6U"))
                                                                {
                                                                    for (int Y = 0; Y < Extend6U.Length; Y++)
                                                                    {
                                                                        string[] arr6U = platform.ToUpper().Split(new string[] { "6U" }, StringSplitOptions.RemoveEmptyEntries);

                                                                        DataRow row6U = dtConfig_PN.NewRow();
                                                                        row6U["ODM"] = strODM;
                                                                        row6U["Platform"] = arr6U[0].Trim();
                                                                        row6U["Size"] = strHierarchy;
                                                                        row6U["Extension"] = Extend6U[Y];  //6U, R6U, RPL6U, RT6U
                                                                        row6U["filesheet"] = sheet_name;
                                                                        row6U["filepath"] = filepath;
                                                                        row6U["rowId"] = rowId;
                                                                        row6U["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                        row6U["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                        row6U["SA_PN"] = xlSheet.Cells[j, ConfigforPNCol].Value.ToString().Trim();

                                                                        //Cost by Size
                                                                        foreach (List<string> innerlist in SizeColList)
                                                                        {
                                                                            string size = innerlist[0].ToString().Trim();
                                                                            int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                            string hierarchy = row6U["Size"].ToString();

                                                                            if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                            {
                                                                                hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                                if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                                {
                                                                                    if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                    {
                                                                                        row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                            else //for 13, 14, 15
                                                                            {
                                                                                if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                                {
                                                                                    if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                    {
                                                                                        row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        dtConfig_PN.Rows.Add(row6U);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    //004: Multi Size (Check Size column 'V')
                                                    if (xlSheet.Cells[j, cateCol + 1].Value != null)
                                                    {
                                                        for (int Z = 0; Z < dvHierarchy.ToTable().Rows.Count; Z++)
                                                        {
                                                            string keyword = xlSheet.Cells[j, cateCol - 1].Value.ToString();
                                                            int quotemark = keyword.ToCharArray().Count(c => c == '"');

                                                            if (quotemark == 0)
                                                            {
                                                                DataRow row = dtOption.NewRow();
                                                                row["ODM"] = strODM;
                                                                row["Platform"] = platform;
                                                                row["Size"] = "";
                                                                row["Extension"] = "";  //6U, R6U, RPL6U
                                                                row["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                                row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                row["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                                row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                row["rowId"] = rowId;
                                                                row["Rule_1"] = "NULL";
                                                                row["Rule_2"] = "";
                                                                row["Rule_3"] = "";
                                                                row["filesheet"] = sheet_name;
                                                                row["filepath"] = filepath;
                                                                dtOption.Rows.Add(row);
                                                            }
                                                            else if (quotemark == 1 || quotemark == 3 || quotemark == 5 || quotemark > 6)
                                                            {
                                                                CPCT.Rows[i]["Error_flag"] = 1;
                                                                CPCT.Rows[i]["ErrorMsg"] = "The number of the quotation mark is invalid for the attached rule on row " + j + ".";
                                                                goto End;
                                                            }
                                                            else if (quotemark <= 6)
                                                            {
                                                                string[] arrHierarchy = dvHierarchy.ToTable().Rows[Z]["Platform"].ToString().Trim().Split(' ');
                                                                string strSzie = arrHierarchy[0].ToString().ToUpper().Replace(platform.ToUpper(), string.Empty);

                                                                DataRow row = dtOption.NewRow();
                                                                row["ODM"] = strODM;
                                                                row["Platform"] = platform;
                                                                row["Size"] = strSzie;
                                                                row["Extension"] = "";
                                                                row["filesheet"] = sheet_name;
                                                                row["filepath"] = filepath;
                                                                row["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                                row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                row["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                                row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                row["rowId"] = rowId;

                                                                string[] arrKeyword = keyword.Split('"');
                                                                if (quotemark == 2)
                                                                {
                                                                    row["Rule_1"] = arrKeyword[1];
                                                                    row["Rule_2"] = "";
                                                                    row["Rule_3"] = "";
                                                                }
                                                                else if (quotemark == 4)
                                                                {
                                                                    row["Rule_1"] = arrKeyword[1];
                                                                    row["Rule_2"] = arrKeyword[3];
                                                                    row["Rule_3"] = "";
                                                                }
                                                                else if (quotemark == 6)
                                                                {
                                                                    row["Rule_1"] = arrKeyword[1];
                                                                    row["Rule_2"] = arrKeyword[3];
                                                                    row["Rule_3"] = arrKeyword[5];
                                                                }

                                                                //Cost by Size
                                                                foreach (List<string> innerlist in SizeColList)
                                                                {
                                                                    string size = innerlist[0].ToString().Trim();
                                                                    int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                    string hierarchy = row["Size"].ToString();

                                                                    if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                    {
                                                                        hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                        if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                        {
                                                                            if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                            {
                                                                                row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                    else //for 13, 14, 15
                                                                    {
                                                                        if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                        {
                                                                            if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                            {
                                                                                row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                dtOption.Rows.Add(row);

                                                                //Extension for 6U
                                                                if (platform.ToUpper().Contains("6U"))
                                                                {
                                                                    for (int Y = 0; Y < Extend6U.Length; Y++)
                                                                    {
                                                                        string[] arr6U = platform.ToUpper().Split(new string[] { "6U" }, StringSplitOptions.RemoveEmptyEntries);

                                                                        DataRow row6U = dtOption.NewRow();
                                                                        row6U["ODM"] = strODM;
                                                                        row6U["Platform"] = arr6U[0].Trim();
                                                                        row6U["Size"] = strSzie;
                                                                        row6U["Extension"] = Extend6U[Y];  //6U, R6U, RPL6U, RT6U
                                                                        row6U["filepath"] = filepath;
                                                                        row6U["filesheet"] = sheet_name;
                                                                        row6U["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                                        row6U["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                        row6U["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                                        row6U["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                        row6U["rowId"] = rowId;

                                                                        if (quotemark == 2)
                                                                        {
                                                                            row6U["Rule_1"] = arrKeyword[1];
                                                                            row6U["Rule_2"] = "";
                                                                            row6U["Rule_3"] = "";
                                                                        }
                                                                        else if (quotemark == 4)
                                                                        {
                                                                            row6U["Rule_1"] = arrKeyword[1];
                                                                            row6U["Rule_2"] = arrKeyword[3];
                                                                            row6U["Rule_3"] = "";
                                                                        }
                                                                        else if (quotemark == 6)
                                                                        {
                                                                            row6U["Rule_1"] = arrKeyword[1];
                                                                            row6U["Rule_2"] = arrKeyword[3];
                                                                            row6U["Rule_3"] = arrKeyword[5];
                                                                        }

                                                                        //Cost by Size
                                                                        foreach (List<string> innerlist in SizeColList)
                                                                        {
                                                                            string size = innerlist[0].ToString().Trim();
                                                                            int sizeCol = Convert.ToInt32(innerlist[1]);
                                                                            string hierarchy = row6U["Size"].ToString();

                                                                            if (hierarchy.ToUpper().Contains("W")) //for workstation
                                                                            {
                                                                                hierarchy = hierarchy.Replace("W", string.Empty).Replace("w", string.Empty);
                                                                                if (size.Contains(hierarchy) && size.ToUpper().Contains("W"))
                                                                                {
                                                                                    if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                    {
                                                                                        row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                            else //for 13, 14, 15
                                                                            {
                                                                                if (size.EndsWith(hierarchy) && !size.ToUpper().StartsWith("W"))
                                                                                {
                                                                                    if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                                    {
                                                                                        row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        dtOption.Rows.Add(row6U);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else //only one size
                                                    {
                                                        string keyword = xlSheet.Cells[j, cateCol - 1].Value.ToString();
                                                        int quotemark = keyword.ToCharArray().Count(c => c == '"');

                                                        if (quotemark == 0)
                                                        {
                                                            DataRow row = dtOption.NewRow();
                                                            row["ODM"] = strODM;
                                                            row["Platform"] = platform;
                                                            row["Size"] = "";
                                                            row["Extension"] = "";  //6U, R6U, RPL6U, RT6U
                                                            row["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                            row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                            row["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                            row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                            row["rowId"] = rowId;
                                                            row["Rule_1"] = "NULL";
                                                            row["Rule_2"] = "";
                                                            row["Rule_3"] = "";
                                                            row["filesheet"] = sheet_name;
                                                            row["filepath"] = filepath;
                                                            dtOption.Rows.Add(row);
                                                        }
                                                        else if (quotemark == 1 || quotemark == 3 || quotemark == 5 || quotemark > 6)
                                                        {
                                                            CPCT.Rows[i]["Error_flag"] = 1;
                                                            CPCT.Rows[i]["ErrorMsg"] = "The number of the quotation mark is invalid for the attached rule on row " + j + ".";
                                                            goto End;
                                                        }
                                                        else if (quotemark <= 6)
                                                        {
                                                            DataRow row = dtOption.NewRow();
                                                            row["ODM"] = strODM;
                                                            row["Platform"] = platform;
                                                            row["Size"] = "";
                                                            row["Extension"] = "";  //6U, R6U, RPL6U, RT6U
                                                            row["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                            row["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                            row["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                            row["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                            row["rowId"] = rowId;
                                                            row["filesheet"] = sheet_name;
                                                            row["filepath"] = filepath;

                                                            string[] arrKeyword = keyword.Split('"');
                                                            if (quotemark == 2)
                                                            {
                                                                row["Rule_1"] = arrKeyword[1];
                                                                row["Rule_2"] = "";
                                                                row["Rule_3"] = "";
                                                            }
                                                            else if (quotemark == 4)
                                                            {
                                                                row["Rule_1"] = arrKeyword[1];
                                                                row["Rule_2"] = arrKeyword[3];
                                                                row["Rule_3"] = "";
                                                            }
                                                            else if (quotemark == 6)
                                                            {
                                                                row["Rule_1"] = arrKeyword[1];
                                                                row["Rule_2"] = arrKeyword[3];
                                                                row["Rule_3"] = arrKeyword[5];
                                                            }

                                                            for (int X = SizeColList.Count - 1; X >= 0; X--)
                                                            {
                                                                int sizeCol = Convert.ToInt32(SizeColList[X][1]);
                                                                if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                {
                                                                    row["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                    break;
                                                                }
                                                            }
                                                            dtOption.Rows.Add(row);

                                                            //Extension for 6U
                                                            if (platform.ToUpper().Contains("6U"))
                                                            {
                                                                for (int Y = 0; Y < Extend6U.Length; Y++)
                                                                {
                                                                    string[] arr6U = platform.ToUpper().Split(new string[] { "6U" }, StringSplitOptions.RemoveEmptyEntries);

                                                                    DataRow row6U = dtOption.NewRow();
                                                                    row6U["Platform"] = arr6U[0].Trim();
                                                                    row6U["ODM"] = strODM;
                                                                    row6U["filepath"] = filepath;
                                                                    row6U["filesheet"] = sheet_name;
                                                                    row6U["Config_Option"] = xlSheet.Cells[j, initCol].Value.ToString().Replace("  ", "").Trim();
                                                                    row6U["Config_Rule"] = xlSheet.Cells[j, cateCol - 1].Value.ToString().Trim();
                                                                    row6U["Config_Num"] = xlSheet.Cells[j, initCol + 1].Value?.ToString().ToUpper().Trim();
                                                                    row6U["AV_Category"] = xlSheet.Cells[j, cateCol].Value.ToString().Trim();
                                                                    row6U["rowId"] = rowId;
                                                                    row6U["Size"] = "";
                                                                    row6U["Extension"] = Extend6U[Y];  //6U, R6U, RPL6U, RT6U

                                                                    if (quotemark == 2)
                                                                    {
                                                                        row6U["Rule_1"] = arrKeyword[1];
                                                                        row6U["Rule_2"] = "";
                                                                        row6U["Rule_3"] = "";
                                                                    }
                                                                    else if (quotemark == 4)
                                                                    {
                                                                        row6U["Rule_1"] = arrKeyword[1];
                                                                        row6U["Rule_2"] = arrKeyword[3];
                                                                        row6U["Rule_3"] = "";
                                                                    }
                                                                    else if (quotemark == 6)
                                                                    {
                                                                        row6U["Rule_1"] = arrKeyword[1];
                                                                        row6U["Rule_2"] = arrKeyword[3];
                                                                        row6U["Rule_3"] = arrKeyword[5];
                                                                    }

                                                                    for (int X = SizeColList.Count - 1; X >= 0; X--)
                                                                    {
                                                                        int sizeCol = Convert.ToInt32(SizeColList[X][1]);
                                                                        if (xlSheet.Cells[j, sizeCol].Value != null && xlSheet.Cells[j, sizeCol].Value?.ToString() != "0" && Regex.IsMatch(xlSheet.Cells[j, sizeCol].Value?.ToString(), @"^-?\d+")) //#NULL: -2146826265
                                                                        {
                                                                            row6U["Cost"] = xlSheet.Cells[j, sizeCol].Value;
                                                                            break;
                                                                        }
                                                                    }
                                                                    dtOption.Rows.Add(row6U);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            rowId++;
                                        }
                                    }
                                }

                                if (numSheet == 0)
                                {
                                    CPCT.Rows[i]["Error_flag"] = 1;
                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot find Summary sheet in CPC.";
                                }

                            End:
                                //ep.Save();
                                //ep.Dispose();

                                fs.Close();
                                fs.Dispose();
                            }
                        }
                    }
                    catch (FileNotFoundException)
                    {
                        CPCT.Rows[i]["Error_flag"] = 1;
                        CPCT.Rows[i]["ErrorMsg"] = "The CPC file was occupied by someone when tool was running.";
                    }
                }
            }
            return Tuple.Create(dtOption, dtRCTO, dtRCTO_PN, dtConfig_PN, dtExclude);
        }

        static void UploadOption(SqlConnection conn, DataTable CPCT, DataTable dtAdder, DataTable dtRCTO, DataTable dtRCTO_PN, DataTable dtConfig_PN, DataTable dtExclude, int weekday, string ver)
        {
            //01: Delete Historical Data
            string strClear = "DELETE From [dbo].[Eevee_File_History] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_BUSA] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_Option_NB] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_Option_NB_RCTO] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_Option_NB_RCTO_PN] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_Option_NB_Config_PN] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 " +
                              "DELETE From [dbo].[Eevee_Upload_Report] WHERE DATEDIFF(MONTH, [UploadTime], GETUTCDATE())>6 ";
            SqlCommand cmdClear = new SqlCommand(strClear, conn);
            cmdClear.CommandTimeout = 0;
            cmdClear.ExecuteNonQuery();

            //02: Upload RCTO, RCTO_PN, Adder
            string strTemp =
                "DROP TABLE IF EXISTS ##tempOption " +
                "CREATE TABLE ##tempOption([ODM] CHAR(30), [Platform] CHAR(60), [Size] CHAR(10), [Extension] CHAR(20), [Config_Option] CHAR(400), [Config_Rule] CHAR(200), [Config_Num] CHAR(30), [Rule_1] CHAR(60), [Rule_2] CHAR(60), [Rule_3] CHAR(60), " +
                "[AV_Category] CHAR(80), [Cost] FLOAT, [rowId] INT, [filepath] CHAR(300), [filesheet] CHAR(50), [UploadTime] SMALLDATETIME, [version] CHAR(20)) " +
                "DROP TABLE IF EXISTS ##tempExclude " +
                "CREATE TABLE ##tempExclude([Platform] CHAR(60), [Exclude] CHAR(60), [filepath] CHAR(300), [filesheet] CHAR(50))";
            SqlCommand cmdTemp = new SqlCommand(strTemp, conn);
            cmdTemp.ExecuteNonQuery();

            SqlBulkCopy sbcExclude = new SqlBulkCopy(conn);
            sbcExclude.DestinationTableName = "[dbo].[##tempExclude]";
            sbcExclude.ColumnMappings.Add("Platform", "Platform");
            sbcExclude.ColumnMappings.Add("Exclude", "Exclude");
            sbcExclude.ColumnMappings.Add("filepath", "filepath");
            sbcExclude.ColumnMappings.Add("filesheet", "filesheet");
            sbcExclude.BulkCopyTimeout = 0;
            sbcExclude.WriteToServer(dtExclude);
            sbcExclude.Close();

            //dtExclude.Clear();
            //dtExclude.Dispose();  //retain for clear SA result

            SqlBulkCopy sbcRCTO = new SqlBulkCopy(conn);
            sbcRCTO.DestinationTableName = "[dbo].[Eevee_Option_NB_RCTO]";
            sbcRCTO.ColumnMappings.Add("Platform", "Platform");
            sbcRCTO.ColumnMappings.Add("ODM", "ODM");
            sbcRCTO.ColumnMappings.Add("RCTO_Site", "RCTO_Site_Code");
            sbcRCTO.ColumnMappings.Add("filesheet", "filesheet");
            sbcRCTO.ColumnMappings.Add("filepath", "filepath");
            sbcRCTO.ColumnMappings.Add("UploadTime", "UploadTime");
            sbcRCTO.ColumnMappings.Add("version", "version");
            sbcRCTO.BulkCopyTimeout = 0;
            sbcRCTO.WriteToServer(dtRCTO);
            sbcRCTO.Close();

            dtRCTO.Clear();
            dtRCTO.Dispose();

            SqlBulkCopy sbcRCTO_PN = new SqlBulkCopy(conn);
            sbcRCTO_PN.DestinationTableName = "[dbo].[Eevee_Option_NB_RCTO_PN]";
            sbcRCTO_PN.ColumnMappings.Add("ODM", "ODM");
            sbcRCTO_PN.ColumnMappings.Add("Platform", "Platform");
            sbcRCTO_PN.ColumnMappings.Add("Config_Rule", "Config_Rule");
            sbcRCTO_PN.ColumnMappings.Add("PMatrix", "PMatrix");
            sbcRCTO_PN.ColumnMappings.Add("rowId", "rowId");
            sbcRCTO_PN.ColumnMappings.Add("RCTO_Site", "RCTO_Site_Code");
            sbcRCTO_PN.ColumnMappings.Add("filesheet", "filesheet");
            sbcRCTO_PN.ColumnMappings.Add("filepath", "filepath");
            sbcRCTO_PN.ColumnMappings.Add("UploadTime", "UploadTime");
            sbcRCTO_PN.ColumnMappings.Add("version", "version");
            sbcRCTO_PN.BulkCopyTimeout = 0;
            sbcRCTO_PN.WriteToServer(dtRCTO_PN);
            sbcRCTO_PN.Close();

            dtRCTO_PN.Clear();
            dtRCTO_PN.Dispose();

            //SqlBulkCopy sbcConfig_PN = new SqlBulkCopy(conn);
            //sbcConfig_PN.DestinationTableName = "[dbo].[Eevee_Option_NB_Config_PN]";
            //sbcConfig_PN.ColumnMappings.Add("ODM", "ODM");
            //sbcConfig_PN.ColumnMappings.Add("Platform", "Platform");
            //sbcConfig_PN.ColumnMappings.Add("Size", "Size");
            //sbcConfig_PN.ColumnMappings.Add("Extension", "Extension");
            //sbcConfig_PN.ColumnMappings.Add("rowId", "rowId");
            //sbcConfig_PN.ColumnMappings.Add("Config_Rule", "Config_Rule");
            //sbcConfig_PN.ColumnMappings.Add("AV_Category", "AV_Category");
            //sbcConfig_PN.ColumnMappings.Add("SA_PN", "SA_PN");
            //sbcConfig_PN.ColumnMappings.Add("Cost", "Cost");
            //sbcConfig_PN.ColumnMappings.Add("filesheet", "filesheet");
            //sbcConfig_PN.ColumnMappings.Add("filepath", "filepath");
            //sbcConfig_PN.ColumnMappings.Add("UploadTime", "UploadTime");
            //sbcConfig_PN.ColumnMappings.Add("version", "version");
            //sbcConfig_PN.BulkCopyTimeout = 0;
            //sbcConfig_PN.WriteToServer(dtConfig_PN);
            //sbcConfig_PN.Close();

            //dtConfig_PN.Clear();
            //dtConfig_PN.Dispose();

            SqlBulkCopy sbcAdder = new SqlBulkCopy(conn);
            sbcAdder.DestinationTableName = "[dbo].[##tempOption]";
            sbcAdder.ColumnMappings.Add("Platform", "Platform");
            sbcAdder.ColumnMappings.Add("ODM", "ODM");
            sbcAdder.ColumnMappings.Add("Extension", "Extension");
            sbcAdder.ColumnMappings.Add("Size", "Size");
            sbcAdder.ColumnMappings.Add("Config_Option", "Config_Option");
            sbcAdder.ColumnMappings.Add("Config_Rule", "Config_Rule");
            sbcAdder.ColumnMappings.Add("Config_Num", "Config_Num");
            sbcAdder.ColumnMappings.Add("Rule_1", "Rule_1");
            sbcAdder.ColumnMappings.Add("Rule_2", "Rule_2");
            sbcAdder.ColumnMappings.Add("Rule_3", "Rule_3");
            sbcAdder.ColumnMappings.Add("AV_Category", "AV_Category");
            sbcAdder.ColumnMappings.Add("Cost", "Cost");
            sbcAdder.ColumnMappings.Add("rowId", "rowId");
            sbcAdder.ColumnMappings.Add("filepath", "filepath");
            sbcAdder.ColumnMappings.Add("filesheet", "filesheet");
            sbcAdder.ColumnMappings.Add("UploadTime", "UploadTime");
            sbcAdder.ColumnMappings.Add("version", "version");
            sbcAdder.BulkCopyTimeout = 0;
            sbcAdder.WriteToServer(dtAdder);
            sbcAdder.Close();

            //03: Update vendor code for RCTO, RCTO_PN
            string strVendorcode =
                "UPDATE t1 SET [Vendor_Code]=t2.[Vendor_Code] FROM [dbo].[Eevee_Option_NB_RCTO] t1, [dbo].[BB8_ODM_Lookup] t2 WHERE t1.[version]='" + ver + "' and t2.[Note]='RCTO' and t2.[BU]='NB' and t1.[RCTO_Site_Code]=t2.[Site_Code] and t1.[ODM] LIKE '%'+ RTRIM(t2.[ODM]) +'%' " +
                "UPDATE t1 SET [Vendor_Code]=t2.[Vendor_Code] FROM [dbo].[Eevee_Option_NB_RCTO_PN] t1, [dbo].[BB8_ODM_Lookup] t2 WHERE t1.[version]='" + ver + "' and t2.[Note]='RCTO' and t2.[BU]='NB' and t1.[RCTO_Site_Code]=t2.[Site_Code] and t1.[ODM] LIKE '%'+ RTRIM(t2.[ODM]) +'%'"; //and t2.[Note]='RCTO'
            SqlCommand cmdVendorcode = new SqlCommand(strVendorcode, conn);
            cmdVendorcode.CommandTimeout = 0;
            cmdVendorcode.ExecuteNonQuery();

            //04: Find Adder PN
            string strFindSA =
                "INSERT INTO [dbo].[Eevee_Option_NB]([ODM], [Platform], [Extension], [Size], [Config_Option], [Config_Rule], [Config_Num], [Rule_1], [Rule_2], [Rule_3], [AV_Category], [Cost], [rowId], [PMatrix], [SA_PN], [SA_Desc], [SA_Cate], [filepath], [filesheet], [UploadTime], [version]) " +
                "SELECT Distinct t1.[ODM], t1.[Platform], [Extension], [Size], [Config_Option], [Config_Rule], [Config_Num], [Rule_1], [Rule_2], [Rule_3], [AV_Category], t1.[Cost], [rowId], t2.[Family], t2.[SA_Number], t2.[SA_Description], t2.[Category], [filepath], [filesheet], t1.[UploadTime], [version] FROM [dbo].[##tempOption] t1 " +
                "LEFT JOIN [dbo].[PMatrix_AVSA_latest] t2 ON t2.[Family] LIKE '' + RTRIM(t1.[Platform]) + '%' and t2.[Family] LIKE '%'+RTRIM(t1.[Extension])+'%' and RTRIM(t2.[Family]) LIKE '%' + LTRIM(RTRIM(t1.[Size])) + '' " +
                "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_1]) + '%' collate Chinese_PRC_CS_AI " +
                "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_2]) + '%' collate Chinese_PRC_CS_AI " +
                "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_3]) + '%' collate Chinese_PRC_CS_AI " +
                "and t2.[Category] LIKE '%' + RTRIM(t1.[AV_Category]) + '%' and LTRIM(t2.[SA_Description]) NOT LIKE 'GNRC%' and LTRIM(t2.[SA_Description]) NOT LIKE 'SSA%' " +
                "" +
                "DELETE FROM [dbo].[Eevee_Option_NB] WHERE [version]='" + ver + "' and [Size]<>'' and [Size] NOT LIKE '%W%' and [PMatrix] is not NULL and SUBSTRING([PMatrix], LEN([Platform])+1, LEN([PMatrix])) LIKE '%W%' " +
                "DELETE FROM [dbo].[Eevee_Option_NB] FROM [dbo].[Eevee_Option_NB] as A, [dbo].[##tempExclude] as B WHERE A.[version]='" + ver + "' and A.[filepath]=B.[filepath] and A.[filesheet]=B.[filesheet] and A.[PMatrix]=B.[Exclude]";
            SqlCommand cmdFindSA = new SqlCommand(strFindSA, conn);
            cmdFindSA.CommandTimeout = 0;
            cmdFindSA.ExecuteNonQuery();

            //04: Config for PN
            string strConfigPN = "UPDATE t1 SET t1.[PMatrix]=t2.[Family], t1.[SA_Desc]=t2.[SA_Description], t1.[SA_Cate]=t2.[Category] FROM [dbo].[Eevee_Option_NB_Config_PN] t1, " +
                "(SELECT distinct A.*, B.[Family], B.[SA_Number], B.[SA_Description], B.[Category] FROM [dbo].[Eevee_Option_NB_Config_PN] A " +
                "LEFT JOIN [dbo].[PMatrix_AVSA_latest] as B on B.[SA_Number]=A.[SA_PN] and B.[Family] LIKE '%'+RTRIM(A.[Platform])+'%' and B.[Family] LIKE '%'+RTRIM(A.[Size])+'%' and B.[Family] LIKE '%'+RTRIM(A.[Extension])+'%' and B.[Category] LIKE '%'+RTRIM(A.[AV_Category])+'%' " +
                "WHERE A.[version]='" + ver + "' and B.[Family] is not null) t2 WHERE t1.[int_Config_for_PN_NB_ID]=t2.[int_Config_for_PN_NB_ID] " +
                "" +
                "DELETE FROM [dbo].[Eevee_Option_NB_Config_PN] WHERE [version]='" + ver + "' and [PMatrix] is null " +
                "DELETE FROM [dbo].[Eevee_Option_NB_Config_PN] WHERE [version]='" + ver + "' and [PMatrix] NOT LIKE '%'+RTRIM([Platform])+'%' and [PMatrix] NOT LIKE '%'+RTRIM([Size])+'%' and [PMatrix] NOT LIKE '%'+RTRIM([Extension])+'%'";
            SqlCommand cmdConfigPN = new SqlCommand(strConfigPN, conn);
            cmdConfigPN.CommandTimeout = 0;
            cmdConfigPN.ExecuteNonQuery();

            string strUpload =
               "INSERT INTO [dbo].[Eevee_Upload_Report]([PN], [Sitecode], [Vendorcode], [Cost], [EffDateFrom], [EffDateTo], [BU], [Type], [UploadTime], [filepath], [version]) " +
               "SELECT Distinct [SA_PN], [Site_Code], [Vendor_Code], SUM(t1.[Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
               "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], [filepath], [version], [Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
               "LEFT JOIN [dbo].[BB8_ODM_Lookup] as B on B.[ODM]=A.[ODM] " +
               "WHERE [SA_PN] is not NULL and [version]='" + ver + "' and [Note]='IDS' and [BU]='NB') as t1 GROUP BY [SA_PN], [Site_Code], [Vendor_Code], [filepath], [version] " +
               "UNION " +
               "SELECT [SA_PN], [RCTO_Site_Code], [Vendor_Code], SUM([Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
               "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], A.[filepath], A.[version], [RCTO_Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
               "LEFT JOIN [dbo].[Eevee_Option_NB_RCTO] as B on B.[filesheet]=A.[filesheet] and B.[filepath]=A.[filepath] and B.[version]=A.[version] " +
               "WHERE [SA_PN] is not NULL and [RCTO_Site_Code] is not NULL and A.[version]='" + ver + "') as t1 GROUP BY [SA_PN], [RCTO_Site_Code], [Vendor_Code], [filepath], [version] " +
               "UNION " +
               "SELECT [SA_PN], [RCTO_Site_Code], [Vendor_Code], SUM([Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
               "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], A.[filepath], A.[version], [RCTO_Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
               "LEFT JOIN [dbo].[Eevee_Option_NB_RCTO_PN] as B on B.[filesheet]=A.[filesheet] and B.[filepath]=A.[filepath] and B.[version]=A.[version] and B.[rowId]=A.[rowId] and B.[PMatrix]=A.[PMatrix] " +
               "WHERE [SA_PN] is not NULL and [RCTO_Site_Code] is not NULL and A.[version]='" + ver + "') as t1 GROUP BY [SA_PN], [RCTO_Site_Code], [Vendor_Code], [filepath], [version] " +
               "Order by [filepath], [SA_PN], [Site_Code]";
            SqlCommand cmdUpload = new SqlCommand(strUpload, conn);
            cmdUpload.CommandTimeout = 0;
            cmdUpload.ExecuteNonQuery();


            //**For Quote Tool: BUSA HP Price
            if (weekday == 2 && DateTime.UtcNow.Day > 12 && DateTime.UtcNow.Day <= 20)
            {
                string strQPlus_BU = "DELETE FROM [dbo].[QPlus_HP_Price] WHERE [Quote_Type]='Final' and [EffDate]='" + DateTime.UtcNow.ToString("yyyy-MM-01") + "' " +
                    "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct [BUSA], [PMatrix_SA_Desc], 'BUSA', [Total_Cost], [ODM], [PMatrix_Family], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', [filepath], 'BU SA', GETUTCDATE(), 'Eevee', [version], [PMatrix_Category] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [BUSA], [ODM], [PMatrix_Family] ORDER BY [int_BUSA_ID] desc) as [srank] FROM [dbo].[Eevee_BUSA] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "' and [PMatrix_Family] is not null) A " +
                    "WHERE [srank]=1 order by [BUSA], [ODM]";
                SqlCommand cmdQPlus_BU = new SqlCommand(strQPlus_BU, conn);
                cmdQPlus_BU.CommandTimeout = 0;
                cmdQPlus_BU.ExecuteNonQuery();

                string strQPlus_OP = "INSERT INTO [dbo].[QPlus_HP_Price]([DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Config]) " +
                    "SELECT Distinct 'OP', [Cost], [ODM], [PMatrix], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', [filepath], [filesheet], GETUTCDATE(), 'Eevee', [version], [Config_Num] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [ODM], [PMatrix], [Config_Num] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "' and [SA_PN] is not null) A " +
                    "WHERE [srank]=1 order by [filepath], [Config_Num]";
                SqlCommand cmdQPlus_OP = new SqlCommand(strQPlus_OP, conn);
                cmdQPlus_OP.CommandTimeout = 0;
                cmdQPlus_OP.ExecuteNonQuery();

                // + program matrix category???
                string strQPlus_Adder = "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct A.[SA_PN], A.[SA_Desc], 'OP', [TTL], A.[ODM], A.[PMatrix], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', A.[filepath], A.[filesheet], GETUTCDATE(), 'Eevee', A.[version], A.[SA_Cate] FROM [dbo].[Eevee_Option_NB] A " +
                    "INNER JOIN (SELECT Distinct [SA_PN], [ODM], SUM(t1.[Cost]) as [TTL], [filepath], [version] FROM (SELECT * FROM (SELECT Distinct [SA_PN], [Cost], [ODM], [Config_Option], [filepath], [version], ROW_NUMBER() OVER (PARTITION BY [SA_PN], [ODM], [Config_Option] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] " +
                    "WHERE [SA_PN] is not NULL and [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "') tb WHERE [srank]=1) as t1 " +
                    "GROUP BY [SA_PN], [ODM], [filepath], [version]) B on B.[SA_PN]=A.[SA_PN] and B.[ODM]=A.[ODM] and B.[filepath]=A.[filepath] and B.[version]=A.[version] order by [PMatrix], [SA_PN], [ODM]";
                SqlCommand cmdQPlus_Adder = new SqlCommand(strQPlus_Adder, conn);
                cmdQPlus_Adder.CommandTimeout = 0;
                cmdQPlus_Adder.ExecuteNonQuery();
            }
            else if (weekday == 2 && DateTime.UtcNow.Day > 20 && DateTime.UtcNow.Day <= 30)
            {
                string strQPlus_BU = "DELETE FROM [dbo].[QPlus_HP_Price] WHERE [Quote_Type]='Initial' and [EffDate]='" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "' " +
                    "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct [BUSA], [PMatrix_SA_Desc], 'BUSA', [Total_Cost], [ODM], [PMatrix_Family], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', [filepath], 'BU SA', GETUTCDATE(), 'Eevee', [version], [PMatrix_Category] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [BUSA], [ODM], [PMatrix_Family] ORDER BY [int_BUSA_ID] desc) as [srank] FROM [dbo].[Eevee_BUSA] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "' and [PMatrix_Family] is not null) A " +
                    "WHERE [srank]=1 order by [BUSA], [ODM]";
                SqlCommand cmdQPlus_BU = new SqlCommand(strQPlus_BU, conn);
                cmdQPlus_BU.CommandTimeout = 0;
                cmdQPlus_BU.ExecuteNonQuery();

                string strQPlus_OP = "INSERT INTO [dbo].[QPlus_HP_Price]([DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Config]) " +
                    "SELECT Distinct 'OP', [Cost], [ODM], [PMatrix], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', [filepath], [filesheet], GETUTCDATE(), 'Eevee', [version], [Config_Num] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [ODM], [PMatrix], [Config_Num] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "' and [SA_PN] is not null) A " +
                    "WHERE [srank]=1 order by [filepath], [Config_Num]";
                SqlCommand cmdQPlus_OP = new SqlCommand(strQPlus_OP, conn);
                cmdQPlus_OP.CommandTimeout = 0;
                cmdQPlus_OP.ExecuteNonQuery();

                // + program matrix category???
                string strQPlus_Adder = "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct A.[SA_PN], A.[SA_Desc], 'OP', [TTL], A.[ODM], A.[PMatrix], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', A.[filepath], A.[filesheet], GETUTCDATE(), 'Eevee', A.[version], A.[SA_Cate] FROM [dbo].[Eevee_Option_NB] A " +
                    "INNER JOIN (SELECT Distinct [SA_PN], [ODM], SUM(t1.[Cost]) as [TTL], [filepath], [version] FROM (SELECT * FROM (SELECT Distinct [SA_PN], [Cost], [ODM], [Config_Option], [filepath], [version], ROW_NUMBER() OVER (PARTITION BY [SA_PN], [ODM], [Config_Option] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] " +
                    "WHERE [SA_PN] is not NULL and [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "') tb WHERE [srank]=1) as t1 " +
                    "GROUP BY [SA_PN], [ODM], [filepath], [version]) B on B.[SA_PN]=A.[SA_PN] and B.[ODM]=A.[ODM] and B.[filepath]=A.[filepath] and B.[version]=A.[version] order by [PMatrix], [SA_PN], [ODM]";
                SqlCommand cmdQPlus_Adder = new SqlCommand(strQPlus_Adder, conn);
                cmdQPlus_Adder.CommandTimeout = 0;
                cmdQPlus_Adder.ExecuteNonQuery();
            }
        }

        static DataTable WriteResult(SqlConnection conn, DataTable CPCT, DataTable dtHierarchy, DataTable dtExclude, string ver)
        {
            DataView dvExclude = new DataView(dtExclude);
            DataView dvPMatrix = new DataView(dtHierarchy);
          
            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                if (CPCT.Rows[i]["Error_flag"].ToString() == "")
                {
                    string filename = CPCT.Rows[i]["filename"].ToString().Trim();
                    string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                    string filesheet = "";  //summary sheet name

                    FileInfo fi = new FileInfo(filepath);
                    using (ExcelPackage ep = new ExcelPackage(fi))
                    {
                        for (int j = 0; j < ep.Workbook.Worksheets.Count; j++)
                        {
                            ExcelWorksheet xlSheet = ep.Workbook.Worksheets[j];
                            string sheetname = xlSheet.Name.ToString().Trim();
                            if (sheetname.StartsWith("Summary"))
                            {
                                if (xlSheet != null)
                                {
                                    filesheet = xlSheet.Name.ToString().Trim();  //summary sheet name

                                    SqlCommand cmdSAPivot = new SqlCommand("Eevee_SA_Pivot", conn);
                                    cmdSAPivot.CommandType = CommandType.StoredProcedure;
                                    cmdSAPivot.Parameters.Add("@filepath", SqlDbType.NVarChar);
                                    cmdSAPivot.Parameters.Add("@filesheet", SqlDbType.NVarChar);
                                    cmdSAPivot.Parameters.Add("@version", SqlDbType.NVarChar);
                                    cmdSAPivot.Parameters["@filepath"].Value = filepath;
                                    cmdSAPivot.Parameters["@filesheet"].Value = filesheet;
                                    cmdSAPivot.Parameters["@version"].Value = ver;
                                    cmdSAPivot.CommandTimeout = 0;
                                    cmdSAPivot.ExecuteNonQuery();

                                    DataTable dtSAPivot = new DataTable();
                                    SqlDataAdapter sdaSAPivot = new SqlDataAdapter(cmdSAPivot);
                                    sdaSAPivot.Fill(dtSAPivot);

                                    int CateCol = 0, CateRow = 0;
                                    for (int X = 1; X <= xlSheet.Dimension.End.Column; X++)
                                    {
                                        for (int Y = 1; Y <= xlSheet.Dimension.End.Row; Y++)
                                        {
                                            if (xlSheet.Cells[Y, X].Text.ToString().Trim().Contains("Category"))  //xlSheet.Cells[Y, X].Text.ToString() != "" && 
                                            {
                                                CateCol = X;
                                                CateRow = Y;
                                                goto WriteSA;
                                            }
                                        }
                                    }

                                WriteSA:
                                    dvExclude.RowFilter = string.Empty;
                                    dvExclude.RowFilter = "filepath LIKE '" + filepath + "%'";
                                    dvPMatrix.RowFilter = string.Empty;
                                    dvPMatrix.RowFilter = "filepath LIKE '" + filepath + "%'";

                                    int excludeNum = dvExclude.ToTable().Rows.Count;
                                    int pmatrixNum = dvPMatrix.ToTable().Rows.Count;

                                    int Plus = 0;
                                    if (CPCT.Rows[i]["RCTO_PN_flag"].ToString() == "1")
                                    {
                                        Plus++;
                                    }

                                    //001: Clear last SA Result
                                    for (int X = CateRow + 1; X <= xlSheet.Dimension.End.Row; X++)
                                    {
                                        for (int Y = 0; Y < pmatrixNum - excludeNum + 1; Y++)
                                        {
                                            if (xlSheet.Cells[X, CateCol - 1].Text.ToString() != "")  //xlSheet.Cells[X, CateCol - 1].Value != null
                                            {
                                                try
                                                {
                                                    xlSheet.Cells[X - 1, CateCol + (2 + Plus) + Y].Value = "";
                                                    xlSheet.Cells[X, CateCol + (2 + Plus) + Y].Value = "";
                                                }
                                                catch (Exception)
                                                {
                                                    CPCT.Rows[i]["Error_flag"] = 1;
                                                    CPCT.Rows[i]["ErrorMsg"] = "Cannot fill in SA results. Please unprotect the cells behind size column.";
                                                    goto End;
                                                }
                                            }
                                        }
                                    }

                                    //002: Write SA Result
                                    if (dtSAPivot.Columns.Count == 0)
                                    {
                                        CPCT.Rows[i]["Null_flag"] = 1;
                                        for (int X = CateRow + 1; X <= xlSheet.Dimension.End.Row; X++)
                                        {
                                            try
                                            {
                                                xlSheet.Cells[CateRow, CateCol + (2 + Plus)].Value = "";
                                                if (xlSheet.Cells[X, CateCol].Text.ToString() != "" && xlSheet.Cells[X, CateCol - 1].Text.ToString() != "")  //xlSheet.Cells[X, CateCol].Value != null && xlSheet.Cells[X, CateCol - 1].Value != null
                                                {
                                                    string category = xlSheet.Cells[X, CateCol].Text.ToString();
                                                    if (!category.ToLower().Contains("base") && !category.ToUpper().Contains("BU") && !category.ToLower().Contains("skip"))
                                                    {
                                                        xlSheet.Cells[X, CateCol + (2 + Plus)].Value = "NULL";
                                                    }
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                CPCT.Rows[i]["Error_flag"] = 1;
                                                CPCT.Rows[i]["ErrorMsg"] = "Cannot fill in SA results. Please unprotect the cells behind size column.";
                                                goto End;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int X = 3; X < dtSAPivot.Columns.Count; X++)  //00:filepath 01: file_sheet 02: rowId
                                        {
                                            try
                                            {
                                                xlSheet.Cells[CateRow, CateCol + (2 + Plus)].Value = dtSAPivot.Columns[X].ColumnName.ToString().Trim() + " (" + DateTime.UtcNow.ToString("M/d") + ")";
                                                for (int Y = 0; Y < dtSAPivot.Rows.Count; Y++)
                                                {
                                                    int row = Convert.ToInt32(dtSAPivot.Rows[Y]["rowId"]);
                                                    if (dtSAPivot.Rows[Y][X].ToString() == "")
                                                    {
                                                        xlSheet.Cells[CateRow + row, CateCol + (2 + Plus)].Value = "NULL";
                                                        CPCT.Rows[i]["Null_flag"] = 1;
                                                    }
                                                    else
                                                    {
                                                        xlSheet.Cells[CateRow + row, CateCol + (2 + Plus)].Value = dtSAPivot.Rows[Y][X].ToString().Trim();
                                                    }
                                                }
                                                CateCol++;
                                            }
                                            catch (Exception)
                                            {
                                                CPCT.Rows[i]["Error_flag"] = 1;
                                                CPCT.Rows[i]["ErrorMsg"] = "Cannot fill in SA results. Please unprotect the cells behind size column.";
                                                goto End;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    CPCT.Rows[i]["Error_flag"] = 1;
                                    CPCT.Rows[i]["ErrorMsg"] = "The CPC file was occupied by someone when tool was running.";
                                }
                            }
                            else if (xlSheet.Name.Contains("Option") && xlSheet.Name.Contains("SA"))   //003: Delete Records
                            {
                                ep.Workbook.Worksheets.Delete(xlSheet);
                                j = j - 1;
                            }
                        }

                        DataTable dtSA = new DataTable();
                        string strSA = "SELECT Distinct [SA_PN] as [SA PN], [SA_Desc] as [SA Desc], [Config_Option] as [Config Option], [Config_Rule] as [Config Rule], [AV_Category] as [Category], [Cost], [PMatrix] as [Program Matrix], [ODM] From [dbo].[Eevee_Option_NB] " +
                            "WHERE [filepath] LIKE '" + filepath + "%' and [SA_PN] is not NULL and [version]='" + ver + "' order by [ODM], [SA PN], [Config Option]";
                        SqlDataAdapter sdaSA = new SqlDataAdapter(strSA, conn);
                        sdaSA.SelectCommand.CommandTimeout = 0;
                        sdaSA.Fill(dtSA);

                        DataTable dtSUM = new DataTable();
                        string strSUM = "SELECT [SA_PN] as [SA PN], [SA_Desc] as [SA Desc], SUM([Cost]) AS [Cost], [PMatrix] as [Program Matrix], [ODM] FROM " +
                            "(SELECT Distinct [SA_PN], [SA_Desc], [Cost], [rowId], [AV_Category], [PMatrix], [ODM] FROM [dbo].[Eevee_Option_NB] " +
                            "WHERE [filepath] LIKE '" + filepath + "%' and [SA_PN] is not NULL and [version]='" + ver + "') as t1 GROUP BY [SA_PN], [SA_Desc], [PMatrix], [ODM] order by [SA PN], [ODM]";
                        SqlDataAdapter sdaSUM = new SqlDataAdapter(strSUM, conn);
                        sdaSUM.SelectCommand.CommandTimeout = 0;
                        sdaSUM.Fill(dtSUM);

                        try
                        {
                            //004: Add worksheet: OptionSA
                            ExcelWorksheet SheetSA = ep.Workbook.Worksheets.Add("OptionSA_" + DateTime.UtcNow.ToString("MMdd"));
                            ep.Workbook.Worksheets.MoveAfter(SheetSA.Name, filesheet);
                            for (int X = 0; X < dtSA.Columns.Count; X++)
                            {
                                SheetSA.Cells[1, X + 1].Value = dtSA.Columns[X].ColumnName.ToString().Trim();
                                for (int Y = 0; Y < dtSA.Rows.Count; Y++)
                                {
                                    if (X == 5)
                                    {
                                        SheetSA.Cells[Y + 2, X + 1].Value = Convert.ToDouble(dtSA.Rows[Y][X].ToString().Trim());
                                    }
                                    else
                                    {
                                        SheetSA.Cells[Y + 2, X + 1].Value = dtSA.Rows[Y][X].ToString().Trim();
                                    }
                                }
                            }

                            ExcelRange titleSA = SheetSA.Cells[1, 1, 1, 8];
                            titleSA.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                            titleSA.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                            titleSA.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            titleSA.Style.Font.Bold = true;
                            SheetSA.Columns[1, 8].AutoFit();

                            //005: Add worksheet: OptionSA_SUM
                            ExcelWorksheet SheetSUM = ep.Workbook.Worksheets.Add("OptionSA_SUM_" + DateTime.UtcNow.ToString("MMdd"));
                            ep.Workbook.Worksheets.MoveAfter(SheetSUM.Name, SheetSA.Name);
                            for (int X = 0; X < dtSUM.Columns.Count; X++)
                            {
                                SheetSUM.Cells[1, X + 1].Value = dtSUM.Columns[X].ColumnName.ToString().Trim();
                                for (int Y = 0; Y < dtSUM.Rows.Count; Y++)
                                {
                                    if (X == 2)
                                    {
                                        SheetSUM.Cells[Y + 2, X + 1].Value = Convert.ToDouble(dtSUM.Rows[Y][X].ToString().Trim());
                                    }
                                    else
                                    {
                                        SheetSUM.Cells[Y + 2, X + 1].Value = dtSUM.Rows[Y][X].ToString().Trim();
                                    }
                                }
                            }

                            ExcelRange titleSUM = SheetSUM.Cells[1, 1, 1, 5];
                            titleSUM.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                            titleSUM.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                            titleSUM.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            titleSUM.Style.Font.Bold = true;
                            SheetSUM.Columns[1, 5].AutoFit();   //titleSUM.AutoFitColumns();
                        }
                        catch (ArgumentException)
                        {
                            string skip = "";
                        }

                    End:
                        try
                        {
                            ep.Save();
                            ep.Dispose();
                        }
                        catch (InvalidOperationException)
                        {
                            ep.Dispose();
                        }
                    }
                }
            }

            //06: Update CPCT Info - Owner/ RCTO flag/ Error flag/ Null flag
            string strInfo = "DROP TABLE IF EXISTS ##tempInfo " +
                "CREATE TABLE ##tempInfo ([Owner] CHAR(80), [RCTO_flag] INT, [RCTO_PN_flag] INT, [Null_flag] INT, [Error_flag] INT, [ErrorMsg] CHAR(150), [version] CHAR(20), [filepath] CHAR(200))";
            SqlCommand cmdInfo = new SqlCommand(strInfo, conn);
            cmdInfo.ExecuteNonQuery();

            SqlBulkCopy sbcInfo = new SqlBulkCopy(conn);
            sbcInfo.DestinationTableName = "##tempInfo";
            sbcInfo.ColumnMappings.Add("Owner", "Owner");
            sbcInfo.ColumnMappings.Add("RCTO_flag", "RCTO_flag");
            sbcInfo.ColumnMappings.Add("RCTO_PN_flag", "RCTO_PN_flag");
            sbcInfo.ColumnMappings.Add("Null_flag", "Null_flag");
            sbcInfo.ColumnMappings.Add("Error_flag", "Error_flag");
            sbcInfo.ColumnMappings.Add("ErrorMsg", "ErrorMsg");
            sbcInfo.ColumnMappings.Add("version", "version");
            sbcInfo.ColumnMappings.Add("filepath", "filepath");
            sbcInfo.WriteToServer(CPCT);
            sbcInfo.Close();

            string strUpdate = "UPDATE t1 SET [Owner]=t2.[Owner], [RCTO_flag]=t2.[RCTO_flag], [RCTO_PN_flag]=t2.[RCTO_PN_flag], [Null_flag]=t2.[Null_flag], [Error_flag]=t2.[Error_flag], [ErrorMsg]=t2.[ErrorMsg] FROM [dbo].[Eevee_File_History] t1, [dbo].[##tempInfo] t2 WHERE t1.[filepath]=t2.[filepath] and t1.[version]=t2.[version]";
            SqlCommand cmdUpdate = new SqlCommand(strUpdate, conn);
            cmdUpdate.ExecuteNonQuery();

            DataTable dtOwner = new DataTable();
            string strOwner = "SELECT Distinct [Owner] FROM [dbo].[Eevee_File_History] Where [version]='" + ver + "' and [Owner] is not null and [Owner]<>''";
            SqlDataAdapter sdaOwner = new SqlDataAdapter(strOwner, conn);
            sdaOwner.Fill(dtOwner);

            return dtOwner;
        }



        static List<string> SendEmail(DataTable CPCT, DataTable dtOwner)
        {
            List<string> allOwner = new List<string>();
            string itembNB = null, itemcNB = null, itemIEC = null, itemWistron = null, itemQuanta = null, itemCompal = null, itemHuaqin = null, itemBYD = null, itemPega = null;
            string missOwner = null;

            SmtpClient client = new SmtpClient("smtp3.hp.com", 25);
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("karon.tang@hp.com", "lfv.yac-14");

            //01: CPCT Email to PBM
            for (int X = 0; X < dtOwner.Rows.Count; X++)
            {
                string NotNone = null, None = null;
                string theOwner = dtOwner.Rows[X]["Owner"].ToString().Trim();

                for (int Y = 0; Y < CPCT.Rows.Count; Y++)
                {
                    if (CPCT.Rows[Y]["Owner"].ToString().Trim() == theOwner && CPCT.Rows[Y]["ErrorMsg"].ToString() == "")
                    {
                        string filename = CPCT.Rows[Y]["FileName"].ToString().Trim();
                        string filepath = CPCT.Rows[Y]["FilePath"].ToString().Trim();
                        string platform = CPCT.Rows[Y]["Platform"].ToString().Trim();

                        if (CPCT.Rows[Y]["Null_flag"].ToString() == "")
                        {
                            NotNone += "<a href=\"" + filepath + "\"> Link to " + platform + "</a> <b>:</b> " + filename + "</br>";
                        }
                        else
                        {
                            None += "<a href=\"" + filepath + "\"> Link to " + platform + "</a> <b>:</b> " + filename + "</br>";
                        }
                    }
                }
                if (NotNone != null || None != null)
                {
                    MailMessage mailToPBM = new MailMessage();
                    mailToPBM.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                    try
                    {
                        mailToPBM.To.Add(theOwner);
                        mailToPBM.Bcc.Add("psgopbmisapp@hp.com");
                        mailToPBM.IsBodyHtml = true;
                        mailToPBM.Subject = "Eevee Tool: CPC tracker was run by Eevee tool.";
                        mailToPBM.SubjectEncoding = Encoding.UTF8;

                        if (NotNone != null)
                        {
                            mailToPBM.Body += "The completed CPC file as below:</br>" + NotNone + "</br>";
                        }
                        if (None != null)
                        {
                            mailToPBM.Body += "The CPC file with NULL cost adder as below:</br>" + None + "</br>";
                        }
                        mailToPBM.Body += "Please review your CPC results. Thank you.";
                        mailToPBM.BodyEncoding = Encoding.UTF8;
                        client.Send(mailToPBM);

                        allOwner.Add(theOwner);
                    }
                    catch (SmtpFailedRecipientException)   //(FormatException)
                    {
                        for (int Z = 0; Z < CPCT.Rows.Count; Z++)
                        {
                            if (CPCT.Rows[Z]["Owner"].ToString().Trim() == theOwner)
                            {
                                CPCT.Rows[Z]["Owner"] = null;
                                CPCT.Rows[Z]["Error_flag"] = 1;
                                CPCT.Rows[Z]["ErrorMsg"] = "Cannot recognize owner's email address.";
                            }
                        }
                        continue;
                    }
                }
            }

            //02: Error Msg + Miss Owner
            for (int X = 0; X < CPCT.Rows.Count; X++)
            {
                string filename = CPCT.Rows[X]["filename"].ToString().Trim();
                string filepath = CPCT.Rows[X]["filepath"].ToString().Trim();
                string platform = CPCT.Rows[X]["Platform"].ToString().Trim();
                string tracking = CPCT.Rows[X]["trackingname"].ToString().Trim();
                string errormsg = CPCT.Rows[X]["ErrorMsg"].ToString().Trim();

                if (errormsg != "")
                {
                    if (filepath.Contains("Compal"))
                    {
                        itemCompal += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Quanta"))
                    {
                        itemQuanta += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Wistron"))
                    {
                        itemWistron += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Inventec"))
                    {
                        itemIEC += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Huaqin"))
                    {
                        itemHuaqin += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("BYD"))
                    {
                        itemBYD += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Pegatron"))
                    {
                        itemPega += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("bNB"))
                    {
                        itembNB += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                    else if (filepath.Contains("Team Site") && filepath.Contains("Trackers"))
                    {
                        itemcNB += "<a href=\"" + filepath + "\">Link to " + tracking + "</a> &lt;Error Message: " + errormsg + "></br>";
                    }
                }
                else if (CPCT.Rows[X]["Owner"].ToString() == "")
                {
                    missOwner += "<a href=\"" + filepath + "\"> Link to " + platform + "</a> <b>:</b> " + filename + "</br>";
                }
            }

            if (itembNB != null || itemcNB != null || itemCompal != null || itemIEC != null || itemQuanta != null || itemWistron != null || itemHuaqin != null || itemBYD != null || itemPega != null)
            {
                MailMessage mailError = new MailMessage();
                mailError.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
                foreach (var theOwner in allOwner)
                {
                    mailError.To.Add(theOwner);
                }
                mailError.Bcc.Add("psgopbmisapp@hp.com");
                mailError.Bcc.Add("karon.tang@hp.com");

                mailError.IsBodyHtml = true;
                mailError.Priority = MailPriority.High;
                mailError.Subject = "Eevee Tool: CPC tracker cannot be run by Eevee tool.";
                mailError.SubjectEncoding = Encoding.UTF8;
                mailError.Body = "The CPC files with issue as below:</br></br>";
                if (itembNB != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>NPI bNB</span></br>" + itembNB + "</br>";
                }
                if (itemcNB != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>NPI cNB </span></br>" + itemcNB + "</br>";
                }
                if (itemCompal != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Compal</span></br>" + itemCompal + "</br>";
                }
                if (itemIEC != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining IEC</span></br>" + itemIEC + "</br>";
                }
                if (itemQuanta != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Quanta</span></br>" + itemQuanta + "</br>";
                }
                if (itemWistron != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Wistron</span></br>" + itemWistron + "</br>";
                }
                if (itemHuaqin != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Huaqin</span></br>" + itemHuaqin + "</br>";
                }
                if (itemBYD != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Huaqin</span></br>" + itemBYD + "</br>";
                }
                if (itemPega != null)
                {
                    mailError.Body += "<span style='background:Yellow; mso-highlight:Yellow'>Sustaining Pegatron</span></br>" + itemPega + "</br>";
                }
                mailError.Body += "Please refer to the error message and revise your CPC files. Thank you.";
                mailError.BodyEncoding = Encoding.UTF8;
                client.Send(mailError);
            }
            else
            {
                MailMessage mailSuccess = new MailMessage();
                mailSuccess.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
                mailSuccess.Bcc.Add("psgopbmisapp@hp.com");
                mailSuccess.Bcc.Add("karon.tang@hp.com");
                mailSuccess.IsBodyHtml = true;
                mailSuccess.Subject = "Eevee Tool: CPC tracker have been run successfully by Eevee tool.";
                mailSuccess.SubjectEncoding = Encoding.UTF8;
                mailSuccess.Body = "CPC tracker have been run successfully by Eevee tool. Thank you.";
                mailSuccess.BodyEncoding = Encoding.UTF8;
                client.Send(mailSuccess);
            }

            //04: Miss Owner
            if (missOwner != null)
            {
                MailMessage mailMissing = new MailMessage();
                mailMissing.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                foreach (var theOwner in allOwner)
                {
                    mailMissing.To.Add(theOwner);
                }
                mailMissing.Bcc.Add("psgopbmisapp@hp.com");
                mailMissing.IsBodyHtml = true;
                mailMissing.Subject = "Eevee Tool: The contact information is missing or cannot be recognized.";
                mailMissing.SubjectEncoding = Encoding.UTF8;
                mailMissing.Body = "The CPC file without contact information as below:</br></br>" + missOwner + "</br>Please revise your contact information for above CPC files. Thank you.";
                mailMissing.BodyEncoding = Encoding.UTF8;
                client.Send(mailMissing);
            }

            return allOwner;
        }

        static string[] Report(SqlConnection conn, DataTable CPCT, int weekday, string ver, string[] outcome)
        {
            //00: error message
            string[] validation = { "", "", "", "", "", "" };  //BUSA, Adder, PIR creation & condition, Cost Difference, Negative Cost, ODM Mgmt

            //01: Update CPC Info
            string strInfo = "DROP TABLE IF EXISTS ##tempInfo CREATE TABLE ##tempInfo ([Owner] CHAR(80), [Error_flag] INT, [ErrorMsg] CHAR(150), [version] CHAR(20), [filepath] CHAR(200))";
            SqlCommand cmdInfo = new SqlCommand(strInfo, conn);
            cmdInfo.ExecuteNonQuery();

            SqlBulkCopy sbcInfo = new SqlBulkCopy(conn);
            sbcInfo.DestinationTableName = "##tempInfo";
            sbcInfo.ColumnMappings.Add("Owner", "Owner");
            sbcInfo.ColumnMappings.Add("Error_flag", "Error_flag");
            sbcInfo.ColumnMappings.Add("ErrorMsg", "ErrorMsg");
            sbcInfo.ColumnMappings.Add("version", "version");
            sbcInfo.ColumnMappings.Add("filepath", "filepath");
            sbcInfo.WriteToServer(CPCT);
            sbcInfo.Close();

            string strUpdate = "UPDATE t1 SET [Owner]=t2.[Owner], [Error_flag]=t2.[Error_flag], [ErrorMsg]=t2.[ErrorMsg] FROM [dbo].[Eevee_File_History] t1, [dbo].[##tempInfo] t2 WHERE t1.[filepath]=t2.[filepath] and t1.[version]=t2.[version]";
            SqlCommand cmdUpdate = new SqlCommand(strUpdate, conn);
            cmdUpdate.ExecuteNonQuery();

            //*001: Cost Difference Report
            DataTable dtCostDiff = new DataTable();
            string strCostDiff = "SELECT Distinct t2.*, t3.[Owner] FROM (SELECT *, ROW_NUMBER() OVER(PARTITION BY [PN], [Sitecode], [Vendorcode], [Type] ORDER BY [Cost] desc) AS [srow] FROM " +
                "(SELECT Distinct [PN], [Sitecode], [Vendorcode], CONVERT(DECIMAL(10, 6), [Cost]) as [Cost], [version], IIF([Type]<>'Adder', 'SA', 'Adder') as [Type] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') A ) t1 " +
                "LEFT JOIN (SELECT *, IIF([Type]<>'Adder', 'SA', 'Adder') as [_Type_] FROM [dbo].[Eevee_Upload_Report]) as t2 on t2.[PN]=t1.[PN] and t2.[Sitecode]=t1.[Sitecode] and t2.[Vendorcode]=t1.[Vendorcode] and t2.[version]=t1.[version] and t2.[_Type_]=t1.[Type] " +
                "LEFT JOIN [dbo].[Eevee_File_History] as t3 on t3.[filepath]=t2.[filepath] and t3.[version]=t2.[version] " +
                "WHERE [srow]>1 order by t2.[PN], t2.[Sitecode], t2.[Vendorcode], [Cost]";
            SqlDataAdapter sdaCostDiff = new SqlDataAdapter(strCostDiff, conn);
            sdaCostDiff.SelectCommand.CommandTimeout = 0;
            sdaCostDiff.Fill(dtCostDiff);
            if (dtCostDiff.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[7]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "PN";
                    xlSheet.Cells[1, 2].Value = "Site Code";
                    xlSheet.Cells[1, 3].Value = "Vendor Code";
                    xlSheet.Cells[1, 4].Value = "Cost";
                    xlSheet.Cells[1, 5].Value = "Type";
                    xlSheet.Cells[1, 6].Value = "Filename";
                    xlSheet.Cells[1, 7].Value = "Owner";

                    for (int X = 0; X < dtCostDiff.Rows.Count; X++)
                    {
                        string[] path = dtCostDiff.Rows[X]["filepath"].ToString().Trim().Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries);
                        string owner = dtCostDiff.Rows[X]["Owner"].ToString().Trim();

                        xlSheet.Cells[X + 2, 1].Value = dtCostDiff.Rows[X]["PN"].ToString().Trim();
                        xlSheet.Cells[X + 2, 2].Value = dtCostDiff.Rows[X]["Sitecode"].ToString().Trim();
                        xlSheet.Cells[X + 2, 3].Value = dtCostDiff.Rows[X]["Vendorcode"].ToString().Trim();
                        xlSheet.Cells[X + 2, 4].Value = Convert.ToDouble(dtCostDiff.Rows[X]["Cost"].ToString().Trim());
                        xlSheet.Cells[X + 2, 5].Value = dtCostDiff.Rows[X]["Type"].ToString().Trim();
                        xlSheet.Cells[X + 2, 6].Value = path[path.Length - 1].Trim();
                        xlSheet.Cells[X + 2, 7].Value = owner;

                        if (!validation[3].Contains(owner))
                        {
                            validation[3] += owner + ";";
                        }
                    }
                    xlSheet.Columns[1, 7].AutoFit();
                    xlSheet.View.FreezePanes(2, 1);

                    ExcelRange title = xlSheet.Cells[1, 1, 1, 7];
                    title.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                    title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                    title.Style.Font.Bold = true;

                    ep.Save();
                    ep.Dispose();
                }
            }

            //*002: Remove cost diff items
            string strDiff = "DELETE FROM [dbo].[Eevee_Upload_Report] WHERE [int_Cost_Upload_ID] IN (SELECT Distinct t2.[int_Cost_Upload_ID] FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode], [Type] ORDER BY [Cost] desc) AS [srow] FROM " +
                "(SELECT Distinct [PN], [Sitecode], [Vendorcode], CONVERT(DECIMAL(10, 6), [Cost]) as [Cost], [version], IIF([Type]<>'Adder', 'SA', 'Adder') as [Type] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') A) t1 " +
                "LEFT JOIN (SELECT *, IIF([Type]<>'Adder', 'SA', 'Adder') as [_Type_] FROM [dbo].[Eevee_Upload_Report]) as t2 on t2.[PN]=t1.[PN] and t2.[Sitecode]=t1.[Sitecode] and t2.[Vendorcode]=t1.[Vendorcode] and t2.[version]=t1.[version] and t2.[_Type_]=t1.[Type] " +
                "LEFT JOIN [dbo].[Eevee_File_History] as t3 on t3.[filepath]=t2.[filepath] and t3.[version]=t2.[version] WHERE [srow]>1)";
            SqlCommand cmdDiff = new SqlCommand(strDiff, conn);
            cmdDiff.ExecuteNonQuery();

            //Wait for 10 seconds
            Thread.Sleep(10000); 

            DataTable dtMLCC = new DataTable();
            string strMLCC = "SELECT *, DATEADD(DAY, -1, DATEADD(MONTH, 3, [DateFrom])) as [DateTo] FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY [DateFrom]) as [srow] FROM [dbo].[BB8_Adder_MLCC] WHERE [DateFrom]>=DATEADD(MONTH, -3, GETUTCDATE())) A";
            SqlDataAdapter sdaMLCC = new SqlDataAdapter(strMLCC, conn);
            sdaMLCC.SelectCommand.CommandTimeout = 0;
            sdaMLCC.Fill(dtMLCC);

            //*003: CCS Upload (New)
            string[,] cost_type = { { "NB", "C" }, { "NB", "F" }, { "NH", "C" } };

            DataTable dtiCost = new DataTable();
            string striCost = "SELECT DISTINCT A.*, B.[Owner] " +
                            "FROM [dbo].[Eevee_Upload_Report] AS A " +
                            "LEFT JOIN (SELECT DISTINCT [filepath], [version], [Owner] FROM [dbo].[Eevee_File_History]) AS B " +
                            "ON A.[filepath] = B.[filepath] AND A.[version] = B.[version] " +
                            "WHERE A.[version] = '" + ver + "' " +
                            "AND A.[Cost] IS NOT NULL AND A.[Vendorcode] IS NOT NULL AND A.[Type] <> 'Adder' " +
                            "ORDER BY A.[filepath], A.[PN], A.[Type] DESC";

            SqlDataAdapter sdaiCost = new SqlDataAdapter(striCost, conn);
            sdaiCost.SelectCommand.CommandTimeout = 0;
            sdaiCost.Fill(dtiCost);

            DataView dviCost = new DataView(dtiCost);

            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                if (CPCT.Rows[i]["ErrorMsg"].ToString() == "")
                {
                    string trackingname = CPCT.Rows[i]["trackingname"].ToString().Trim();
                    string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();

                    dviCost.RowFilter = $"filepath LIKE '{filepath}%'";

                    DataTable filteredTable = dviCost.ToTable();

                    FileInfo fi = new FileInfo($"{outcome[0]}\\{trackingname} -EffDate {ver} Id{100 + i} (New).xlsx");
                    using (ExcelPackage ep = new ExcelPackage(fi))
                    {
                        ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");

                        // Set header
                        string[] headers = {
                            "Site", "PartNo.", "Vendor", "TransactionType", "CostType", "Cost",
                            "EffFromDate", "Currentcy", "Exch. Rate", "Contract No.",
                            "Target Quantity", "Target Value", "Onwer"
                        };
                        for (int h = 0; h < headers.Length; h++)
                            xlSheet.Cells[1, h + 1].Value = headers[h];

                        xlSheet.Columns[1].Style.Numberformat.Format = "@";
                        xlSheet.Columns[3].Style.Numberformat.Format = "@";

                        int srow = 2;

                        foreach (DataRow row in filteredTable.Rows)
                        {
                            string owner = row["Owner"]?.ToString().Trim() ?? "UNKNOWN";

                            for (int y = 0; y < cost_type.GetLength(0); y++)
                            {
                                xlSheet.Cells[srow, 1].Value = row["Sitecode"].ToString().Trim();
                                xlSheet.Cells[srow, 2].Value = row["PN"].ToString().Trim();
                                xlSheet.Cells[srow, 3].Value = row["Vendorcode"].ToString().Trim();
                                xlSheet.Cells[srow, 4].Value = cost_type[y, 0];
                                xlSheet.Cells[srow, 5].Value = cost_type[y, 1];

                                double cost = 0.01;
                                if (row["Cost"] != DBNull.Value && double.TryParse(row["Cost"].ToString(), out double parsedCost))
                                    cost = parsedCost == 0 ? 0.01 : parsedCost;

                                xlSheet.Cells[srow, 6].Value = cost;
                                xlSheet.Cells[srow, 7].Value = DateTime.UtcNow.ToString("MM/dd/yyyy");
                                xlSheet.Cells[srow, 8].Value = "USD";
                                xlSheet.Cells[srow, 9].Value = ""; // Exch. Rate
                                xlSheet.Cells[srow, 10].Value = ""; // Contract No.
                                xlSheet.Cells[srow, 11].Value = ""; // Target Qty
                                xlSheet.Cells[srow, 12].Value = ""; // Target Value
                                xlSheet.Cells[srow, 13].Value = owner;

                                srow++;
                            }
                        }

                        ep.Save();
                    }
                }
            }

            //*004: PIR Creation
            DataTable dtPIRcreate = new DataTable();
            string strPIRcreate = "SELECT Distinct [PN], [Sitecode], [Vendorcode] FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode] ORDER BY [int_Cost_Upload_ID]) AS [srow] FROM [dbo].[Eevee_Upload_Report]) A WHERE [srow]=1 and [Vendorcode] is not null and [version]='" + ver + "'";
            SqlDataAdapter sdaPIRcreate = new SqlDataAdapter(strPIRcreate, conn);
            sdaPIRcreate.Fill(dtPIRcreate);
            if (dtPIRcreate.Rows.Count > 0)
            {
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet[1, 1].Value = "Vendor Number";
                worksheet[1, 2].Value = "Purchase Organization";
                worksheet[1, 3].Value = "Plant";
                worksheet[1, 4].Value = "Material Number";
                worksheet[1, 5].Value = "Info Record Category";
                worksheet[1, 6].Value = "Supplier Material Number";
                worksheet[1, 7].Value = "Vendor Sub-range";
                worksheet[1, 8].Value = "Planned Delivery Time in Days";
                worksheet[1, 9].Value = "Purchasing Group";
                worksheet[1, 10].Value = "Standard Purchase Order Quantity";
                worksheet[1, 11].Value = "Minimum Purchase Order Quantity";
                worksheet[1, 12].Value = "Underdelivery Tolerance Limit";
                worksheet[1, 13].Value = "Overdelivery Tolerance Limit";
                worksheet[1, 14].Value = "Indicator: Unlimited";
                worksheet[1, 15].Value = "Indicator: GR-Based Always use CAPITAL letter \"X\"";
                worksheet[1, 16].Value = "Order Acknowledgment Requirement";
                worksheet[1, 17].Value = "Confirmation Control Key";
                worksheet[1, 18].Value = "Tax on Sales/Purchases Code";
                worksheet[1, 19].Value = "Net Price";
                worksheet[1, 20].Value = "Price unit";
                worksheet[1, 21].Value = "Shipping Instructions";
                worksheet[1, 22].Value = "Period Indicator";
                worksheet[1, 23].Value = "Minimum Remaining";
                worksheet[1, 24].Value = "First Reminder/Expediter";
                worksheet[1, 25].Value = "Second Reminder/Expediter";
                worksheet[1, 26].Value = "Third Reminder/Expediter";
                worksheet[1, 27].Value = "Price Determination";
                worksheet[1, 28].Value = "Incoterms (Part 1)";
                worksheet[1, 29].Value = "Incoterms (Part 2)";
                worksheet[1, 30].Value = "Info record Text1";
                worksheet[1, 31].Value = "Info record Text2";
                worksheet[1, 32].Value = "Info record Text3";
                worksheet[1, 33].Value = "Info record Text4";
                worksheet[1, 34].Value = "Info record Text5";
                worksheet[1, 35].Value = "Purchase order Text1";
                worksheet[1, 36].Value = "Purchase order Text2";
                worksheet[1, 37].Value = "Purchase order Text3";
                worksheet[1, 38].Value = "Purchase order Text4";
                worksheet[1, 39].Value = "Purchase order Text5";
                worksheet[1, 40].Value = "Valid-From Date";
                worksheet[1, 41].Value = "Salesperson Responsible";
                worksheet[1, 42].Value = "Country of Issue of Certificate of Origin";
                worksheet[1, 43].Value = "Currency Key";
                worksheet[1, 44].Value = "Order Price Unit (Purchasing)";
                worksheet[1, 45].Value = "Numerator for Conv. of Order Price Unit into Order Unit";
                worksheet[1, 46].Value = "Denominator for Conv. of Order Price Unit into Order Unit";
                worksheet[1, 47].Value = "Condition Group with Vendor";
                worksheet[1, 48].Value = "Rounding Profile";

                for (int X = 0; X < dtPIRcreate.Rows.Count; X++)
                {
                    worksheet[X + 2, 1].Value = "'" + dtPIRcreate.Rows[X]["Vendorcode"].ToString().Trim().Substring(1, 9);
                    worksheet[X + 2, 2].Value = "'" + dtPIRcreate.Rows[X]["Sitecode"].ToString().Trim();
                    worksheet[X + 2, 3].Value = "WCMP";
                    worksheet[X + 2, 4].Value = dtPIRcreate.Rows[X]["PN"].ToString().Trim();
                    worksheet[X + 2, 5].Value = "0";
                    worksheet[X + 2, 8].Value = "7";
                    worksheet[X + 2, 9].Value = "CXX";
                    worksheet[X + 2, 10].Value = "1";
                    worksheet[X + 2, 15].Value = "X";
                    worksheet[X + 2, 18].Value = "I0";
                    worksheet[X + 2, 27].Value = "5";
                    worksheet[X + 2, 40].Value = DateTime.UtcNow.ToString("yyyyMMdd");
                }
                worksheet.SaveToFile(outcome[2], ",", Encoding.UTF8);
                File.WriteAllText(outcome[2], File.ReadAllText(outcome[2]).Replace("\"", ""));
            }

            //*005: PIR Price Condition 
            DataTable dtPIRprice = new DataTable();
            string strPIRprice =
                "SELECT Distinct t1.[PN], t1.[Sitecode], t1.[Vendorcode], ROUND(t2.[Cost], 4) as [Cost], t2.[EffDateTo], t2.[Type] as [_Type_] FROM (SELECT Distinct B.*, C.[Cost] as [previous], B.[Cost]-C.[Cost] as [Delta] FROM " +
                "(SELECT * FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode] ORDER BY [Cost]) as [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') tb WHERE [srow]=1 and [Cost]>=0) A " +
                "LEFT JOIN (SELECT Distinct [PN], [Sitecode], [Vendorcode], [Cost], [version], [Type], ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode], [Type] ORDER BY [int_Cost_Upload_ID] desc) AS [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') as B on B.[PN]=A.[PN] and B.[Sitecode]=A.[Sitecode] and B.[Vendorcode]=A.[Vendorcode] " +
                "FULL JOIN (SELECT Distinct [PN], [Sitecode], [Vendorcode], [Cost], [version], [Type], ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode], [Type] ORDER BY [int_Cost_Upload_ID] desc) AS [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]<>'" + ver + "') as C on C.[PN]=B.[PN] and C.[Sitecode]=B.[Sitecode] and C.[Vendorcode]=B.[Vendorcode] and C.[Type]=B.[Type] " +
                "WHERE B.[PN] is not null and B.[srow]=1 and (C.[srow]=1 OR C.[srow] is null)) t1 " +
                "LEFT JOIN [dbo].[Eevee_Upload_Report] as t2 on t2.[PN]=t1.[PN] and t2.[Sitecode]=t1.[Sitecode] and t2.[Vendorcode]=t1.[Vendorcode] and t2.[version]=t1.[version] " +
                "WHERE [Delta]<>0 or [Delta] is null";
            SqlDataAdapter sdaPIRprice = new SqlDataAdapter(strPIRprice, conn);
            sdaPIRprice.Fill(dtPIRprice);

            if (dtPIRprice.Rows.Count > 0)
            {
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet[1, 1].Value = "Material Number";
                worksheet[1, 2].Value = "Vendor Number";
                worksheet[1, 3].Value = "Purchase Organization";
                worksheet[1, 4].Value = "Plant";
                worksheet[1, 5].Value = "Info Record Category";
                worksheet[1, 6].Value = "Condition type";
                worksheet[1, 7].Value = "Validity start date";
                worksheet[1, 8].Value = "Validity end date";
                worksheet[1, 9].Value = "Condition amount";
                worksheet[1, 10].Value = "Currency";
                worksheet[1, 11].Value = "Price unit";
                worksheet[1, 12].Value = "Deletion Indicator";

                int rrow = 2;
                for (int X = 0; X < dtPIRprice.Rows.Count; X++)
                {
                    string type = dtPIRprice.Rows[X]["_Type_"].ToString().Trim();
                    if (type == "Adder")
                    {
                        for (int Y = 0; Y < dtMLCC.Rows.Count; Y++)
                        {
                            double mlcc_actual = Convert.ToDouble(dtMLCC.Rows[0]["MLCC"]);
                            double mlcc_fcst = Convert.ToDouble(dtMLCC.Rows[Y]["MLCC"]);

                            worksheet[rrow, 1].Value = dtPIRprice.Rows[X]["PN"].ToString().Trim().Replace("\"", "");
                            worksheet[rrow, 2].Value = "'" + dtPIRprice.Rows[X]["Vendorcode"].ToString().Trim().Substring(1, 9);
                            worksheet[rrow, 3].Value = "'" + dtPIRprice.Rows[X]["Sitecode"].ToString().Trim();
                            worksheet[rrow, 4].Value = "WCMP";
                            worksheet[rrow, 5].Value = "0";
                            worksheet[rrow, 6].Value = "ZAD1";

                            if (Y == 0)
                            {
                                worksheet[rrow, 7].Value = DateTime.UtcNow.ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateTo"]).ToString("yyyyMMdd");
                            }
                            else if (Y == dtMLCC.Rows.Count - 1)
                            {
                                worksheet[rrow, 7].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateFrom"]).ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = "99991231";
                            }
                            else
                            {
                                worksheet[rrow, 7].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateFrom"]).ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateTo"]).ToString("yyyyMMdd");
                            }

                            worksheet[rrow, 9].Value = Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) == 0 ? "1" : ((Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) - mlcc_actual + mlcc_fcst) * 100).ToString();
                            worksheet[rrow, 10].Value = "USD";
                            worksheet[rrow, 11].Value = "100";
                            rrow++;
                        }
                    }
                    else if (type == "BU")
                    {
                        for (int Y = 0; Y < dtMLCC.Rows.Count; Y++)
                        {
                            worksheet[rrow, 1].Value = dtPIRprice.Rows[X]["PN"].ToString().Trim();
                            worksheet[rrow, 2].Value = "'" + dtPIRprice.Rows[X]["Vendorcode"].ToString().Trim().Substring(1, 9);
                            worksheet[rrow, 3].Value = "'" + dtPIRprice.Rows[X]["Sitecode"].ToString().Trim();
                            worksheet[rrow, 4].Value = "WCMP";
                            worksheet[rrow, 5].Value = "0";
                            worksheet[rrow, 6].Value = "PB00";
                            if (Y == 0)
                            {
                                worksheet[rrow, 7].Value = DateTime.UtcNow.ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateTo"]).ToString("yyyyMMdd");
                            }
                            else if (Y == dtMLCC.Rows.Count - 1)
                            {
                                worksheet[rrow, 7].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateFrom"]).ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = "99991231";
                            }
                            else
                            {
                                worksheet[rrow, 7].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateFrom"]).ToString("yyyyMMdd");
                                worksheet[rrow, 8].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateTo"]).ToString("yyyyMMdd");
                            }
                            worksheet[rrow, 9].Value = Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) == 0 ? "1" : (Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) * 100).ToString();
                            worksheet[rrow, 10].Value = "USD";
                            worksheet[rrow, 11].Value = "100";
                            rrow++;
                        }
                    }
                    else
                    {
                        worksheet[rrow, 1].Value = dtPIRprice.Rows[X]["PN"].ToString().Trim();
                        worksheet[rrow, 2].Value = "'" + dtPIRprice.Rows[X]["Vendorcode"].ToString().Trim().Substring(1, 9);
                        worksheet[rrow, 3].Value = "'" + dtPIRprice.Rows[X]["Sitecode"].ToString().Trim();
                        worksheet[rrow, 4].Value = "WCMP";
                        worksheet[rrow, 5].Value = "0";
                        worksheet[rrow, 6].Value = "PB00";
                        worksheet[rrow, 7].Value = DateTime.UtcNow.ToString("yyyyMMdd");
                        worksheet[rrow, 8].Value = Convert.ToDateTime(dtPIRprice.Rows[X]["EffDateTo"]).ToString("yyyyMMdd");
                        worksheet[rrow, 9].Value = Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) == 0 ? "1" : (Convert.ToDouble(dtPIRprice.Rows[X]["Cost"]) * 100).ToString();
                        worksheet[rrow, 10].Value = "USD";
                        worksheet[rrow, 11].Value = "100";
                        rrow++;
                    }
                }

                worksheet.SaveToFile(outcome[3], ",", Encoding.UTF8);
                File.WriteAllText(outcome[3], File.ReadAllText(outcome[3]).Replace("\"", ""));
            }

            //*006: CCS NBFA
            DataTable dtNBFA = new DataTable();
            string strNBFA = "SELECT Distinct * FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "' and [Type]='Adder'";
            SqlDataAdapter sdaNBFA = new SqlDataAdapter(strNBFA, conn);
            sdaNBFA.Fill(dtNBFA);

            if (dtNBFA.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[1]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "Site";
                    xlSheet.Cells[1, 2].Value = "PartNo.";
                    xlSheet.Cells[1, 3].Value = "Vendor";
                    xlSheet.Cells[1, 4].Value = "TransactionType";
                    xlSheet.Cells[1, 5].Value = "CostType";
                    xlSheet.Cells[1, 6].Value = "Cost";
                    xlSheet.Cells[1, 7].Value = "EffFromDate";
                    xlSheet.Cells[1, 8].Value = "Currentcy";
                    xlSheet.Cells[1, 9].Value = "Exch. Rate";
                    xlSheet.Cells[1, 10].Value = "Contract No.";
                    xlSheet.Cells[1, 11].Value = "Target Quantity";
                    xlSheet.Cells[1, 12].Value = "Target Value";

                    xlSheet.Columns[1].Style.Numberformat.Format = "@";
                    xlSheet.Columns[3].Style.Numberformat.Format = "@";

                    int srow = 2;
                    for (int X = 0; X < dtNBFA.Rows.Count; X++)
                    {
                        for (int Y = 0; Y < dtMLCC.Rows.Count; Y++)
                        {
                            double mlcc_actual = Convert.ToDouble(dtMLCC.Rows[0]["MLCC"]);
                            double mlcc_fcst = Convert.ToDouble(dtMLCC.Rows[Y]["MLCC"]);

                            xlSheet.Cells[srow, 1].Value = dtNBFA.Rows[X]["Sitecode"].ToString().Trim();
                            xlSheet.Cells[srow, 2].Value = dtNBFA.Rows[X]["PN"].ToString().Trim();
                            xlSheet.Cells[srow, 3].Value = dtNBFA.Rows[X]["Vendorcode"].ToString().Trim();
                            xlSheet.Cells[srow, 4].Value = "NB";
                            xlSheet.Cells[srow, 5].Value = "FA";
                            xlSheet.Cells[srow, 6].Value = Convert.ToDouble(dtNBFA.Rows[X]["Cost"]) == 0 ? 0.01 : Convert.ToDouble(dtNBFA.Rows[X]["Cost"].ToString()) - mlcc_actual + mlcc_fcst;
                            if (Y == 0)
                            {
                                xlSheet.Cells[srow, 7].Value = DateTime.UtcNow.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                xlSheet.Cells[srow, 7].Value = Convert.ToDateTime(dtMLCC.Rows[Y]["DateFrom"]).ToString("MM/dd/yyyy");
                            }
                            xlSheet.Cells[srow, 8].Value = "USD";
                            srow++;
                        }
                    }
                    ep.Save();
                    ep.Dispose();
                }
            }

            //*007: Option Adder
            DataTable dtOption = new DataTable();
            string strOption = "SELECT Distinct [SA_PN], [SA_Desc], [Config_Option], [Config_Rule], [AV_Category], [Cost], [PMatrix] FROM [dbo].[Eevee_Option_NB] WHERE [SA_PN] is not NULL and [version]='" + ver + "' order by [PMatrix], [SA_PN]";
            SqlDataAdapter sdaOption = new SqlDataAdapter(strOption, conn);
            sdaOption.Fill(dtOption);
            if (dtOption.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[4]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "SA PN";
                    xlSheet.Cells[1, 2].Value = "SA Description";
                    xlSheet.Cells[1, 3].Value = "Config Option";
                    xlSheet.Cells[1, 4].Value = "Config Rule";
                    xlSheet.Cells[1, 5].Value = "Category";
                    xlSheet.Cells[1, 6].Value = "Cost";
                    xlSheet.Cells[1, 7].Value = "Program Matrix";

                    for (int X = 0; X < dtOption.Rows.Count; X++)
                    {
                        xlSheet.Cells[X + 2, 1].Value = dtOption.Rows[X]["SA_PN"].ToString().Trim();
                        xlSheet.Cells[X + 2, 2].Value = dtOption.Rows[X]["SA_Desc"].ToString().Trim();
                        xlSheet.Cells[X + 2, 3].Value = dtOption.Rows[X]["Config_Option"].ToString().Trim();
                        xlSheet.Cells[X + 2, 4].Value = dtOption.Rows[X]["Config_Rule"].ToString().Trim();
                        xlSheet.Cells[X + 2, 5].Value = dtOption.Rows[X]["AV_Category"].ToString().Trim();
                        xlSheet.Cells[X + 2, 6].Value = Convert.ToDouble(dtOption.Rows[X]["Cost"].ToString().Trim());
                        xlSheet.Cells[X + 2, 7].Value = dtOption.Rows[X]["PMatrix"].ToString().Trim();
                    }
                    xlSheet.View.FreezePanes(2, 1);      //xlSheet.Columns[1, 7].AutoFit();

                    ExcelRange title = xlSheet.Cells[1, 1, 1, 7];
                    title.Style.Fill.PatternType =XLS.ExcelFillStyle.Solid;
                    title.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                    title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                    title.Style.Font.Bold = true;

                    ep.Save();
                    ep.Dispose();
                }
            }

            //*008: BUSA
            DataTable dtBUSA = new DataTable();
            string strBUSA = "SELECT Distinct A.[Platform], [BUSA], [Description], IIF(B.[ODM] is null, A.[ODM], B.[ODM]) as [ODM], [Total_Cost] as [Cost], [Owner] FROM [dbo].[Eevee_BUSA] A " +
                "LEFT JOIN [dbo].[Eevee_File_History] as B on A.[filepath]=B.[filepath] and A.[version]=B.[version] " +
                "WHERE A.[version]='" + ver + "' and [BUSA] is not null and CAST([Total_Cost] as float)!= 0 order by [Platform], [BUSA], [ODM]";
            SqlDataAdapter sdaBUSA = new SqlDataAdapter(strBUSA, conn);
            sdaBUSA.Fill(dtBUSA);
            if (dtBUSA.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[5]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "Platform";
                    xlSheet.Cells[1, 2].Value = "BU SA";
                    xlSheet.Cells[1, 3].Value = "SA Desc";
                    xlSheet.Cells[1, 4].Value = "ODM";
                    xlSheet.Cells[1, 5].Value = "Cost";
                    xlSheet.Cells[1, 6].Value = "Owner";

                    for (int X = 0; X < dtBUSA.Rows.Count; X++)
                    {
                        xlSheet.Cells[X + 2, 1].Value = dtBUSA.Rows[X]["Platform"].ToString().Trim();
                        xlSheet.Cells[X + 2, 2].Value = dtBUSA.Rows[X]["BUSA"].ToString().Trim();
                        xlSheet.Cells[X + 2, 3].Value = dtBUSA.Rows[X]["Description"].ToString().Trim();
                        xlSheet.Cells[X + 2, 4].Value = dtBUSA.Rows[X]["ODM"].ToString().Trim();
                        xlSheet.Cells[X + 2, 5].Value = Convert.ToDouble(dtBUSA.Rows[X]["Cost"].ToString().Trim());
                        xlSheet.Cells[X + 2, 6].Value = dtBUSA.Rows[X]["Owner"].ToString().Trim();
                    }
                    xlSheet.Columns[1, 6].AutoFit();
                    xlSheet.View.FreezePanes(2, 1);

                    ExcelRange title = xlSheet.Cells[1, 1, 1, 6];
                    title.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                    title.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                    title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                    title.Style.Font.Bold = true;

                    ep.Save();
                    ep.Dispose();
                }
            }

            //*009: BB8 Adders
            DataTable dtBB8 = new DataTable();
            string strBB8 = "SELECT Distinct [BUSA], [Description], [Total_Cost], [Platform], [ODM], [PL], [AV_Site_Committed], [Adder_MLCC], [Adder_P10], [Adder_Resilient], [Adder_RCTO], [Adder_PassThru], [Total_Adder], [PMatrix_SA_Desc], [PMatrix_Category] FROM [dbo].[Eevee_BUSA] " +
                "WHERE [version]='" + ver + "' order by [Platform], [BUSA]";
            SqlDataAdapter sdaBB8 = new SqlDataAdapter(strBB8, conn);
            sdaBB8.Fill(dtBB8);
            if (dtBB8.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[6]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "BU SA";
                    xlSheet.Cells[1, 2].Value = "Description";
                    xlSheet.Cells[1, 3].Value = "Total Cost";
                    xlSheet.Cells[1, 4].Value = "Platform";
                    xlSheet.Cells[1, 5].Value = "ODM";
                    xlSheet.Cells[1, 6].Value = "PL";
                    xlSheet.Cells[1, 7].Value = "AV Site Committed";
                    xlSheet.Cells[1, 8].Value = "MLCC Adder";
                    xlSheet.Cells[1, 9].Value = "P10 Adder";
                    xlSheet.Cells[1, 10].Value = "Resilient Adder";
                    xlSheet.Cells[1, 11].Value = "RCTO Adder";
                    xlSheet.Cells[1, 12].Value = "Pass Thru Adder";
                    xlSheet.Cells[1, 13].Value = "Total Adder";
                    xlSheet.Cells[1, 14].Value = "SA Desc (PMatrix)";
                    xlSheet.Cells[1, 15].Value = "Category (PMatrix)";

                    for (int X = 0; X < dtBB8.Rows.Count; X++)
                    {
                        xlSheet.Cells[X + 2, 1].Value = dtBB8.Rows[X][0].ToString().Trim();
                        xlSheet.Cells[X + 2, 2].Value = dtBB8.Rows[X][1].ToString().Trim();
                        xlSheet.Cells[X + 2, 3].Value = Convert.ToDouble(dtBB8.Rows[X][2].ToString().Trim());
                        xlSheet.Cells[X + 2, 4].Value = dtBB8.Rows[X][3].ToString().Trim();
                        xlSheet.Cells[X + 2, 5].Value = dtBB8.Rows[X][4].ToString().Trim();
                        xlSheet.Cells[X + 2, 6].Value = dtBB8.Rows[X][5].ToString().Trim();
                        xlSheet.Cells[X + 2, 7].Value = dtBB8.Rows[X][6].ToString().Trim();
                        xlSheet.Cells[X + 2, 8].Value = Convert.ToDouble(dtBB8.Rows[X][7].ToString().Trim());
                        xlSheet.Cells[X + 2, 9].Value = Convert.ToDouble(dtBB8.Rows[X][8].ToString().Trim());
                        xlSheet.Cells[X + 2, 10].Value = Convert.ToDouble(dtBB8.Rows[X][9].ToString().Trim());
                        xlSheet.Cells[X + 2, 11].Value = Convert.ToDouble(dtBB8.Rows[X][10].ToString().Trim());
                        xlSheet.Cells[X + 2, 12].Value = Convert.ToDouble(dtBB8.Rows[X][11].ToString().Trim());
                        xlSheet.Cells[X + 2, 13].Value = Convert.ToDouble(dtBB8.Rows[X][12].ToString().Trim());
                        xlSheet.Cells[X + 2, 14].Value = dtBB8.Rows[X][13].ToString().Trim();
                        xlSheet.Cells[X + 2, 15].Value = dtBB8.Rows[X][14].ToString().Trim();
                    }
                    xlSheet.Columns[1, 15].AutoFit();
                    xlSheet.View.FreezePanes(2, 1);

                    ExcelRange title = xlSheet.Cells[1, 1, 1, 15];               
                    title.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                    title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                    title.Style.Font.Bold = true;

                    ExcelRange header = xlSheet.Cells[1, 8, 1, 13];
                    header.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                    header.Style.Fill.BackgroundColor.SetColor(Color.Orchid);

                    ep.Save();
                    ep.Dispose();
                }
            }

            //*010: Negative Cost
            DataTable dtNegative = new DataTable();
            string strNegative = "SELECT Distinct t1.[PN], t1.[Sitecode], t1.[Vendorcode], t2.[Cost] as [_Cost_], IIF(t2.[Type]<>'Adder', 'SA', 'Adder') as [_Type_], t2.[filepath], t3.[Owner] FROM (SELECT Distinct B.*, C.[Cost] as [previous], B.[Cost]-C.[Cost] as [Delta] FROM " +
                "(SELECT * FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode] ORDER BY [Cost]) as [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') tb WHERE [srow]=1 and [Cost]<0) A " +
                "LEFT JOIN (SELECT Distinct [PN], [Sitecode], [Vendorcode], [Cost], [version], IIF([Type]<>'Adder', 'SA', 'Adder') as [Type], ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode], IIF([Type]<>'Adder', 'SA', 'Adder') ORDER BY [int_Cost_Upload_ID] desc) AS [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]='" + ver + "') as B on B.[PN]=A.[PN] and B.[Sitecode]=A.[Sitecode] and B.[Vendorcode]=A.[Vendorcode] " +
                "FULL JOIN (SELECT Distinct [PN], [Sitecode], [Vendorcode], [Cost], [version], IIF([Type]<>'Adder', 'SA', 'Adder') as [Type], ROW_NUMBER() OVER (PARTITION BY [PN], [Sitecode], [Vendorcode], IIF([Type]<>'Adder', 'SA', 'Adder') ORDER BY [int_Cost_Upload_ID] desc) AS [srow] FROM [dbo].[Eevee_Upload_Report] WHERE [version]<>'" + ver + "') as C on C.[PN]=B.[PN] and C.[Sitecode]=B.[Sitecode] and C.[Vendorcode]=B.[Vendorcode] and C.[Type]=B.[Type] " +
                "WHERE B.[PN] is not null and B.[srow]=1 and (C.[srow]=1 OR C.[srow] is null)) t1 " +
                "LEFT JOIN [dbo].[Eevee_Upload_Report] as t2 on t2.[PN]=t1.[PN] and t2.[Sitecode]=t1.[Sitecode] and t2.[Vendorcode]=t1.[Vendorcode] and t2.[version]=t1.[version] " +
                "LEFT JOIN [dbo].[Eevee_File_History] as t3 on t3.[version]=t2.[version] and t3.[filepath]=t2.[filepath]" +
                "WHERE [Delta]<>0 or [Delta] is null";
            SqlDataAdapter sdaNegative = new SqlDataAdapter(strNegative, conn);
            sdaNegative.Fill(dtNegative);
            if (dtNegative.Rows.Count > 0)
            {
                FileInfo fi = new FileInfo(outcome[8]);
                using (ExcelPackage ep = new ExcelPackage(fi))
                {
                    ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                    xlSheet.Cells[1, 1].Value = "PN";
                    xlSheet.Cells[1, 2].Value = "Site Code";
                    xlSheet.Cells[1, 3].Value = "Vendor Code";
                    xlSheet.Cells[1, 4].Value = "Cost";
                    xlSheet.Cells[1, 5].Value = "Type";
                    xlSheet.Cells[1, 6].Value = "Filename";
                    xlSheet.Cells[1, 7].Value = "Owner";

                    for (int X = 0; X < dtNegative.Rows.Count; X++)
                    {
                        string[] filepath = dtNegative.Rows[X]["filepath"].ToString().Trim().Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                        string owner = dtNegative.Rows[X]["Owner"].ToString().Trim();

                        xlSheet.Cells[X + 2, 1].Value = dtNegative.Rows[X]["PN"].ToString().Trim();
                        xlSheet.Cells[X + 2, 2].Value = dtNegative.Rows[X]["Sitecode"].ToString().Trim();
                        xlSheet.Cells[X + 2, 3].Value = dtNegative.Rows[X]["Vendorcode"].ToString().Trim();
                        xlSheet.Cells[X + 2, 4].Value = Convert.ToDouble(dtNegative.Rows[X]["_Cost_"].ToString().Trim());
                        xlSheet.Cells[X + 2, 5].Value = dtNegative.Rows[X]["_Type_"].ToString().Trim();
                        xlSheet.Cells[X + 2, 6].Value = filepath[filepath.Length - 1];  //filepath[filepath.Length - 2] + " >> " + filepath[filepath.Length - 1];
                        xlSheet.Cells[X + 2, 7].Value = owner;

                        if (!validation[4].Contains(owner))
                        {
                            validation[4] += owner + ";";
                        }
                    }
                    xlSheet.Columns[1, 7].AutoFit();
                    xlSheet.View.FreezePanes(2, 1);

                    ExcelRange title = xlSheet.Cells[1, 1, 1, 7];
                    title.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                    title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                    title.Style.Font.Bold = true;

                    ep.Save();
                    ep.Dispose();
                }
            }

            //*011: ODM Managed Cost (one time a month)
            if (weekday == 2)
            {
                DataTable dtMgmt = new DataTable();
                string strMgmt = "SELECT * FROM [dbo].[Eevee_File_History] WHERE [CurrentMonth]='" + DateTime.UtcNow.ToString("yyyy-MM-01") + "' and [ODM_Mgmt_Report]=1";
                SqlDataAdapter sdaMgmt = new SqlDataAdapter(strMgmt, conn);
                sdaMgmt.Fill(dtMgmt);

                if (dtMgmt.Rows.Count == 0)
                {
                    DataTable dtODMMgmt = new DataTable();
                    string strODMMgmt = "SELECT Distinct A.[Platform], [BUSA], [Description], [ODM_Managed_Cost], [Comment], IIF(B.[ODM] is null, A.[ODM], B.[ODM]) as [ODM] FROM [dbo].[Eevee_BUSA] A " +
                        "LEFT JOIN [dbo].[Eevee_File_History] as B on B.[filepath]=A.[filepath] and B.[version]=A.[version] " +
                        "WHERE A.[version]='" + ver + "' and [BUSA] is not null and [ODM_Managed_Cost]<>0";
                    SqlDataAdapter sdaODMMgmt = new SqlDataAdapter(strODMMgmt, conn);
                    sdaODMMgmt.Fill(dtODMMgmt);

                    FileInfo fi = new FileInfo(outcome[9]);
                    using (ExcelPackage ep = new ExcelPackage(fi))
                    {
                        ExcelWorksheet xlSheet = ep.Workbook.Worksheets.Add("Sheet1");
                        xlSheet.Cells[1, 1].Value = "Platform";
                        xlSheet.Cells[1, 2].Value = "BU SA";
                        xlSheet.Cells[1, 3].Value = "SA Desc";
                        xlSheet.Cells[1, 4].Value = "ODM Managed Cost";
                        xlSheet.Cells[1, 5].Value = "ODM";
                        xlSheet.Cells[1, 6].Value = "Comment";

                        for (int X = 0; X < dtODMMgmt.Rows.Count; X++)
                        {
                            xlSheet.Cells[X + 2, 1].Value = dtODMMgmt.Rows[X]["Platform"].ToString().Trim();
                            xlSheet.Cells[X + 2, 2].Value = dtODMMgmt.Rows[X]["BUSA"].ToString().Trim();
                            xlSheet.Cells[X + 2, 3].Value = dtODMMgmt.Rows[X]["Description"].ToString().Trim();
                            xlSheet.Cells[X + 2, 4].Value = Convert.ToDouble(dtODMMgmt.Rows[X]["ODM_Managed_Cost"].ToString().Trim());
                            xlSheet.Cells[X + 2, 5].Value = dtODMMgmt.Rows[X]["ODM"].ToString().Trim();
                            xlSheet.Cells[X + 2, 6].Value = dtODMMgmt.Rows[X]["Comment"].ToString().Trim();
                        }
                        ExcelRange title = xlSheet.Cells[1, 1, 1, 6];
                        title.Style.HorizontalAlignment = XLS.ExcelHorizontalAlignment.Center;
                        title.Style.Fill.PatternType = XLS.ExcelFillStyle.Solid;
                        title.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                        title.Style.Font.Bold = true;
                        xlSheet.Columns[1, 6].AutoFit();

                        ep.Save();

                        validation[5] = "1";
                    }

                    string sqlMgmt = "UPDATE [dbo].[Eevee_File_History] SET [ODM_Mgmt_Report]=1 WHERE [version]='" + ver + "'";
                    SqlCommand cmdMgmt = new SqlCommand(sqlMgmt, conn);
                    cmdMgmt.ExecuteNonQuery();
                }
            }
      

            //00: Error Message
            for (int X = 0; X < CPCT.Rows.Count; X++)
            {
                string errmsg = CPCT.Rows[X]["ErrorMsg"].ToString().Trim();
                if (errmsg != "" && (errmsg.Contains("Cannot find BUSA sheet.") || errmsg.Contains("The heading of BUSA sheet")))
                {
                    validation[0] += CPCT.Rows[X]["FileName"].ToString().Trim() + "<br>";
                }
                if (errmsg != "")
                {
                    validation[1] += CPCT.Rows[X]["FileName"].ToString().Trim() + "<br>";
                }
            }

            if (dtPIRcreate.Rows.Count > 0 && dtPIRprice.Rows.Count > 0)
            {
                validation[2] = "PIR creation and price condition";
            }
            else if (dtPIRcreate.Rows.Count > 0)
            {
                validation[2] = "PIR creation";
            }
            else if (dtPIRprice.Rows.Count > 0)
            {
                validation[2] = "PIR price condition";
            }

            //if (dtCostDiff.Rows.Count > 0)  //Add Owner Email Address
            //{
            //    validation[3] = "1";
            //}
            //if (dtNegative.Rows.Count > 0)  //Add Owner Email Address
            //{
            //    validation[4] = "1";
            //}
        
            return validation;
        }

        static void SendEmail_icost(SqlConnection conn, DataTable CPCT, string ver, string[] validation, string[] outcome, List<string> allOwner)
        {
            DataTable dtPBM = new DataTable();
            string strPBM = "SELECT Distinct [Owner] FROM [dbo].[Eevee_File_History] WHERE [Error_flag] is null and [Owner] is not NULL and [Owner]<>'' and [version]='" + ver + "'";
            SqlDataAdapter sdaPBM = new SqlDataAdapter(strPBM, conn);
            sdaPBM.Fill(dtPBM);

            DataTable dtDL = new DataTable();
            string strDL = "SELECT * FROM [dbo].[Eevee_User] WHERE [PL]='NB' order by [Email]";
            SqlDataAdapter sdaDL = new SqlDataAdapter(strDL, conn);
            sdaDL.Fill(dtDL);
            DataView dvDL = new DataView(dtDL);

            SmtpClient client = new SmtpClient("smtp3.hp.com", 25);
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("karon.tang@hp.com", "Babyface!2021");

            //01: CCS Upload Report
            for (int X = 0; X < dtPBM.Rows.Count; X++)
            {
                string pbm = dtPBM.Rows[X]["Owner"].ToString().Trim();
                string current = DateTime.UtcNow.ToString("yyyy-MM-dd HHmmss");

                string[] CCSfiles = Directory.GetFiles(outcome[0], "*.xlsx");
                using (FileStream fs = new FileStream(outcome[0] + @"\CCS Upload Report(New)-EffDate " + current + ".zip", FileMode.Create))
                using (ZipArchive arch = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    foreach (var CCS in CCSfiles)
                    {
                        FileInfo fi = new FileInfo(CCS);
                        for (int Y = 0; Y < CPCT.Rows.Count; Y++)
                        {
                            string tracking = CPCT.Rows[Y]["trackingname"].ToString().Trim();
                            string errormsg = CPCT.Rows[Y]["ErrorMsg"].ToString().Trim();
                            if (errormsg == "" && CPCT.Rows[Y]["Owner"].ToString().Trim() == pbm && fi.Name.Contains(tracking) && fi.Name.Contains(ver) && fi.Name.Contains("Id" + (100 + Y)))  //+id for Hook, cannonbal
                            {
                                arch.CreateEntryFromFile(fi.FullName, fi.Name);
                            }
                        }
                    }
                }

                MailMessage MsgCCS = new MailMessage();
                MsgCCS.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                dvDL.RowFilter = string.Empty;
                dvDL.RowFilter = "Type='CCS' and EmailTo=1";
                foreach (DataRow row in dvDL.ToTable().Rows)
                {
                    MsgCCS.To.Add(row["Email"].ToString().Trim());
                }
                MsgCCS.To.Add(pbm);
                MsgCCS.Bcc.Add("psgopbmisapp@hp.com");
                MsgCCS.IsBodyHtml = true;
                MsgCCS.Subject = "Eevee Tool: CCS upload report has completed.";
                MsgCCS.SubjectEncoding = Encoding.UTF8;
                MsgCCS.Body = "Please refer to the attachment. Thank you.";
                MsgCCS.Attachments.Add(new Attachment(outcome[0] + @"\CCS Upload Report(New)-EffDate " + current + ".zip"));
                MsgCCS.BodyEncoding = Encoding.UTF8;
                client.Send(MsgCCS);
            }

            //02: PIR creation, price condition
            if (validation[2] != "")
            {
                dvDL.RowFilter = string.Empty;
                dvDL.RowFilter = "Type='PIR' and EmailTo=1";

                MailMessage MsgPIR = new MailMessage();
                MsgPIR.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
                foreach (DataRow row in dvDL.ToTable().Rows)
                {
                    MsgPIR.To.Add(row["Email"].ToString().Trim());
                }
                MsgPIR.Bcc.Add("psgopbmisapp@hp.com");
                MsgPIR.IsBodyHtml = true;
                MsgPIR.Subject = "Eevee Tool: " + validation[2] + " report have completed.";
                MsgPIR.SubjectEncoding = Encoding.UTF8;
                MsgPIR.Body = "Please refer to the attachments for cost upload. Thank you.";
                if (validation[2].Contains("creation"))
                {
                    MsgPIR.Attachments.Add(new Attachment(outcome[2]));
                }
                if (validation[2].Contains("condition"))
                {
                    MsgPIR.Attachments.Add(new Attachment(outcome[3]));
                }
                MsgPIR.BodyEncoding = Encoding.UTF8;
                client.Send(MsgPIR);
            }

            ////03: CCS NBFA
            dvDL.RowFilter = string.Empty;
            dvDL.RowFilter = "Type='NBFA' and EmailTo=1";

            MailMessage MsgNBFA = new MailMessage();
            MsgNBFA.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
            foreach (DataRow row in dvDL.ToTable().Rows)
            {
                MsgNBFA.To.Add(row["Email"].ToString().Trim());
            }
            MsgNBFA.Bcc.Add("psgopbmisapp@hp.com");
            MsgNBFA.IsBodyHtml = true;
            MsgNBFA.Subject = "Eevee Tool: CCS NBFA upload report has been completed.";
            MsgNBFA.SubjectEncoding = Encoding.UTF8;
            MsgNBFA.Body = "Please refer to CCS NBFA upload report as attached. Thank you.</br>";
            MsgNBFA.BodyEncoding = Encoding.UTF8;
            MsgNBFA.Attachments.Add(new Attachment(outcome[1]));
            client.Send(MsgNBFA);

            //001: Option Adder
            dvDL.RowFilter = string.Empty;
            dvDL.RowFilter = "Type LIKE 'Option%' and EmailTo=1";

            MailMessage MsgOption = new MailMessage();
            MsgOption.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
            foreach (DataRow row in dvDL.ToTable().Rows)
            {
                MsgOption.To.Add(row["Email"].ToString().Trim());
            }
            MsgOption.Bcc.Add("psgopbmisapp@hp.com");
            MsgOption.IsBodyHtml = true;
            MsgOption.Subject = "Eevee Tool: Option adder report has been completed.";
            MsgOption.SubjectEncoding = Encoding.UTF8;
            if (validation[1].ToString() != "")
            {
                MsgOption.Body += "The CPC files with issues as below:<br><br>" + validation[1].ToString() + "<br>";
            }
            MsgOption.Body += "Please refer to the attachment. Thank you.";
            MsgOption.Attachments.Add(new Attachment(outcome[4]));
            MsgOption.BodyEncoding = Encoding.UTF8;
            client.Send(MsgOption);

            //002: BUSA 
            dvDL.RowFilter = string.Empty;
            dvDL.RowFilter = "Type='BUSA' and EmailTo=1";

            MailMessage MsgBUSA = new MailMessage();
            MsgBUSA.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
            foreach (DataRow row in dvDL.ToTable().Rows)
            {
                MsgBUSA.To.Add(row["Email"].ToString().Trim());
            }
            foreach (DataRow row in dtPBM.Rows)
            {
                MsgBUSA.CC.Add(row["Owner"].ToString().Trim());
            }
            MsgBUSA.Bcc.Add("psgopbmisapp@hp.com");
            MsgBUSA.IsBodyHtml = true;
            MsgBUSA.Subject = "Eevee Tool: BUSA consolidated report has completed.";
            MsgBUSA.SubjectEncoding = Encoding.UTF8;
            if (validation[0].ToString() != "")
            {
                MsgBUSA.Body += "The CPC files with issues as below:<br><br>" + validation[0].ToString() + "<br>";
            }
            MsgBUSA.Body += "Please refer to the attachment. Thank you.";
            MsgBUSA.Attachments.Add(new Attachment(outcome[5]));
            MsgBUSA.BodyEncoding = Encoding.UTF8;
            client.Send(MsgBUSA);

            //003: BB8 Adders
            dvDL.RowFilter = string.Empty;
            dvDL.RowFilter = "Type LIKE 'Adder%' and EmailTo=1";

            MailMessage MsgBB8 = new MailMessage();
            MsgBB8.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);
            foreach (DataRow row in dvDL.ToTable().Rows)
            {
                MsgBB8.To.Add(row["Email"].ToString().Trim());
            }
            MsgBB8.Bcc.Add("psgopbmisapp@hp.com");
            MsgBB8.IsBodyHtml = true;
            MsgBB8.Subject = "Eevee Tool: BB8 adders list has completed.";
            MsgBB8.SubjectEncoding = Encoding.UTF8;
            if (validation[1].ToString() != "")
            {
                MsgBB8.Body += "The CPC files with issues as below:<br><br>" + validation[1].ToString() + "<br>";
            }
            MsgBB8.Body += "Please refer to the attachment. Thank you.";
            MsgBB8.Attachments.Add(new Attachment(outcome[6]));
            MsgBB8.BodyEncoding = Encoding.UTF8;
            client.Send(MsgBB8);

            //004: Cost difference
            if (validation[3] != "")
            {
                dvDL.RowFilter = string.Empty;
                dvDL.RowFilter = "Type LIKE '%Diff%' and EmailCC=1";

                MailMessage msgCostDiff = new MailMessage();
                msgCostDiff.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                string[] owners = validation[3].Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var i in owners)
                {
                    msgCostDiff.To.Add(i.ToString().Trim());
                }
                foreach (DataRow row in dvDL.ToTable().Rows)
                {
                    msgCostDiff.CC.Add(row["Email"].ToString().Trim());
                }
                msgCostDiff.Bcc.Add("psgopbmisapp@hp.com");
                msgCostDiff.IsBodyHtml = true;
                msgCostDiff.Priority = MailPriority.High;
                msgCostDiff.Subject = "Eevee Tool: Cost Differences list has been generated.";
                msgCostDiff.SubjectEncoding = Encoding.UTF8;
                msgCostDiff.Body = "Please revise the cost different PNs by referring to the attachment. Thank you.";
                msgCostDiff.Attachments.Add(new Attachment(outcome[7]));
                msgCostDiff.BodyEncoding = Encoding.UTF8;
                client.Send(msgCostDiff);
            }

            //005: Negative Cost
            if (validation[4] != "")
            {
                dvDL.RowFilter = string.Empty;
                dvDL.RowFilter = "[Type] LIKE 'Negative%' and [EmailTo]=1";

                MailMessage msgNegative = new MailMessage();
                msgNegative.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                string[] owners = validation[4].Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var i in owners)
                {
                    msgNegative.To.Add(i.ToString().Trim());
                }
                foreach (DataRow row in dvDL.ToTable().Rows)
                {
                    msgNegative.CC.Add(row["Email"].ToString().Trim());
                }
                msgNegative.Bcc.Add("psgopbmisapp@hp.com");
                msgNegative.IsBodyHtml = true;
                msgNegative.Subject = "Eevee Tool: Negative cost report has been generated.";
                msgNegative.SubjectEncoding = Encoding.UTF8;
                msgNegative.Body = "Please refer to negative cost report as attached. Thank you.";
                msgNegative.BodyEncoding = Encoding.UTF8;
                msgNegative.Attachments.Add(new Attachment(outcome[8]));
                client.Send(msgNegative);
            }

            //006: ODM managed cost => Spending Report
            if (validation[5] != "")
            {
                MailMessage msgODMMgmt = new MailMessage();
                msgODMMgmt.From = new MailAddress("psgopbmisapp@hp.com", "PSO PBM BI Service", Encoding.UTF8);

                dvDL.RowFilter = string.Empty;
                dvDL.RowFilter = "Type LIKE 'Spend%' and EmailTo=1";
                foreach (DataRow row in dvDL.ToTable().Rows)
                {
                    msgODMMgmt.To.Add(row["Email"].ToString().Trim());
                }
                msgODMMgmt.CC.Add("karon.tang@hp.com");
                msgODMMgmt.Bcc.Add("psgopbmisapp@hp.com");

                msgODMMgmt.IsBodyHtml = true;
                msgODMMgmt.Subject = "Eevee Tool: BUSA ODM managed cost report has been generated.";
                msgODMMgmt.SubjectEncoding = Encoding.UTF8;
                if (validation[0].ToString() != "")
                {
                    msgODMMgmt.Body += "The CPC files with issues as below:<br><br>" + validation[0].ToString() + "<br>";
                }
                msgODMMgmt.Body += "Please refer to the attachment. Thank you.";
                msgODMMgmt.Attachments.Add(new Attachment(outcome[9]));
                msgODMMgmt.BodyEncoding = Encoding.UTF8;
                client.Send(msgODMMgmt);
            }

        }





        static void QTool(SqlConnection conn, DataTable CPCT, string ver)
        {
            DataTable dtQToolBUSA = new DataTable();
            dtQToolBUSA.Columns.Add("BUSA");
            dtQToolBUSA.Columns.Add("Total_Cost");
            dtQToolBUSA.Columns.Add("ODM");
            dtQToolBUSA.Columns.Add("filepath");
            dtQToolBUSA.Columns.Add("UploadTime");
            dtQToolBUSA.Columns.Add("version");
            dtQToolBUSA.Columns["UploadTime"].DefaultValue = DateTime.UtcNow;
            dtQToolBUSA.Columns["version"].DefaultValue = ver;

            DateTime current = DateTime.Now.AddMonths(-5);

            for (int i = 0; i < CPCT.Rows.Count; i++)
            {
                string filepath = CPCT.Rows[i]["filepath"].ToString().Trim();
                string filename = CPCT.Rows[i]["filename"].ToString().Trim();
                string filepath_rev = @"\\RAICHU\Karon\QTool\" + filename.Replace(".xlsx", "") + "_QTool_Rev.xlsx";
                //string filepath_rev = filepath.Replace(".xlsx", "") + "_QTool_Rev.xlsx";

                File.Copy(filepath, filepath_rev);

                FileInfo fi = new FileInfo(filepath_rev);
                using (ExcelPackage ep = new ExcelPackage(fi))  //using EPPlus
                {
                    int[] tabindex = { -1, -1 };
                    foreach (ExcelWorksheet varSheet in ep.Workbook.Worksheets)
                    {
                        if (varSheet.Name.ToUpper().Contains("CPC") && varSheet.Name.ToLower().Contains("spec change"))
                        {
                            tabindex[0] = varSheet.Index;
                        }
                        else if (varSheet.Name.ToUpper().Contains("CPC") && varSheet.Name.ToLower().Contains("pricing update"))
                        {
                            tabindex[1] = varSheet.Index;
                        }
                    }

                    for (int j = 0; j < tabindex.Length; j++)
                    {
                        if (tabindex[j] > 0)  //tabindex: SPEC Change >> Pricing Update
                        {
                            //001: Find Keyword
                            ExcelWorksheet xlSheet = ep.Workbook.Worksheets[tabindex[j]];
                            int EffCol = 0, EffRow = 0, DescCol = 0;
                            for (int X = 1; X <= xlSheet.Dimension.End.Row; X++)
                            {
                                for (int Y = 1; Y <= xlSheet.Dimension.End.Column; Y++)
                                {
                                    if (xlSheet.Cells[X, Y].Value != null)
                                    {
                                        string cell = xlSheet.Cells[X, Y].Value.ToString();
                                        if (EffCol == 0 && (cell.ToLower().Contains("effectivity date") || cell.ToLower().Contains("effective")))
                                        {
                                            EffCol = Y;
                                            EffRow = X;
                                        }
                                        else if (DescCol == 0 && cell.ToLower().Contains("description"))
                                        {
                                            DescCol = Y;
                                        }
                                    }

                                    if (EffCol > 0 && DescCol > 0)
                                    {
                                        break;
                                    }
                                }
                            }

                            //002:Iterate
                            int count = 1;
                            while (xlSheet.Cells[EffRow + count, DescCol].Value != null)
                            {
                                string eff = "";
                                DateTime? effDate = null;

                                if (xlSheet.Cells[EffRow + count, EffCol].Value != null)
                                {
                                    try
                                    {
                                        effDate = Convert.ToDateTime(xlSheet.Cells[EffRow + count, EffCol].Text.ToString());
                                    }
                                    catch (FormatException)
                                    {
                                        eff = xlSheet.Cells[EffRow + count, EffCol].Text.ToString();
                                    }

                                    if (eff.ToUpper().Contains("TBD") || eff.ToUpper().Contains("TBC") || (effDate != null && effDate > current))
                                    {
                                        xlSheet.Cells[EffRow + count, EffCol, EffRow + count, DescCol].Style.Fill.BackgroundColor.SetColor(Color.Black);
                                        xlSheet.Cells[EffRow + count, EffCol, EffRow + count, DescCol].Value = " ";
                                    }
                                }
                                count++;
                            }
                        }
                    }
                    ep.Workbook.FullCalcOnLoad = true;
                    ep.SaveAs(filepath_rev);
                    ep.Dispose();
                }

                //BU SA
                using (FileStream fs = new FileStream(filepath_rev, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (XLWorkbook xlBook = new XLWorkbook(fs))
                    {
                        //xlBook.CalculateMode = XLCalculateMode.Auto;
                        //xlBook.FullCalculationOnLoad = true;

                        //Process.Start(new ProcessStartInfo(filepath_rev) { UseShellExecute = true });

                        IXLWorksheet worksheet = xlBook.Worksheet("BU SA");
                        if (worksheet != null)
                        {
                            int SACol = 0, CostCol = 0, ODMCol = 0;
                            for (int Y = 1; Y <= worksheet.LastColumnUsed().ColumnNumber(); Y++)
                            {
                                if (worksheet.Cell(1, Y).Value.ToString() != null)
                                {
                                    string cell = worksheet.Cell(1, Y).Value.ToString();
                                    if (SACol == 0 && cell.ToLower().Contains("level 3"))
                                    {
                                        SACol = Y;
                                    }
                                    else if (CostCol == 0 && cell.ToLower().Contains("base unit cost"))
                                    {
                                        CostCol = Y;
                                    }
                                    else if (ODMCol == 0 && cell.ToUpper().Trim() == "ODM")
                                    {
                                        ODMCol = Y;
                                    }
                                }
                            }

                            if (SACol > 0 && CostCol > 0 && ODMCol > 0)
                            {
                                for (int X = 1; X <= worksheet.LastRowUsed().RowNumber(); X++)
                                {
                                    if (worksheet.Cell(X, SACol).Value.ToString() != "" && worksheet.Cell(X, CostCol).Value.ToString() != null && Regex.IsMatch(worksheet.Cell(X, CostCol).Value.ToString(), @"^-?\d+"))  //+- digital number (more)
                                    {

                                        var tesstt = worksheet.Cell(X, SACol);
                                        var tessst = worksheet.Cell(X, CostCol);

                                        DataRow row = dtQToolBUSA.NewRow();
                                        row["BUSA"] = worksheet.Cell(X, SACol).Value.ToString().Trim();
                                        row["Total_Cost"] = Convert.ToDouble(worksheet.Cell(X, CostCol).Value.ToString().Trim());
                                        row["ODM"] = worksheet.Cell(X, ODMCol).Value.ToString().Trim();
                                        row["filepath"] = filepath_rev;
                                        dtQToolBUSA.Rows.Add(row);
                                    }
                                }
                            }
                        }

                    }

                }




                using (FileStream fs = new FileStream(filepath_rev, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    /***
                    using (ExcelPackage ep = new ExcelPackage(fs))
                    {
                        ExcelWorksheet worksheet = ep.Workbook.Worksheets["BU SA"];
                        if (worksheet != null)
                        {
                            int SACol = 0, CostCol = 0, ODMCol = 0;
                            for (int Y = 1; Y <= worksheet.Dimension.End.Column; Y++)
                            {
                                if (worksheet.Cells[1, Y].Value != null)
                                {
                                    string cell = worksheet.Cells[1, Y].Value.ToString();
                                    if (SACol == 0 && cell.ToLower().Contains("level 3"))
                                    {
                                        SACol = Y;
                                    }
                                    else if (CostCol == 0 && cell.ToLower().Contains("base unit cost"))
                                    {
                                        CostCol = Y;
                                    }
                                    else if (ODMCol == 0 && cell.ToUpper().Trim() == "ODM")
                                    {
                                        ODMCol = Y;
                                    }
                                }
                            }

                            if (SACol > 0 && CostCol > 0 && ODMCol > 0)
                            {
                                for (int X = 1; X <= worksheet.Dimension.End.Row; X++)
                                {
                                    if (worksheet.Cells[X, SACol].Text.ToString() != "" && worksheet.Cells[X, CostCol].Value != null && Regex.IsMatch(worksheet.Cells[X, CostCol].Value?.ToString(), @"^-?\d+"))  //+- digital number (more)
                                    {
                                        var tessst = worksheet.Cells[X, CostCol];

                                        DataRow row = dtQToolBUSA.NewRow();
                                        row["BUSA"] = worksheet.Cells[X, SACol].Text.ToString().Trim();
                                        row["Total_Cost"] = Convert.ToDouble(worksheet.Cells[X, CostCol].Value?.ToString().Trim());
                                        row["ODM"] = worksheet.Cells[X, ODMCol].Text.ToString().Trim();
                                        row["filepath"] = filepath_rev;
                                        dtQToolBUSA.Rows.Add(row);
                                    }
                                }
                            }
                        }
                        ep.Save();
                        ep.Dispose();
                    }
                    ***/
                }
            }

            string strTemp = "DROP TABLE IF EXISTS ##tempBUSA  CREATE TABLE ##tempBUSA([BUSA] CHAR(10), [Total_Cost] FLOAT, [ODM] CHAR(30), [filepath] CHAR(300), [UploadTime] SMALLDATETIME, [version] CHAR(20))";
            SqlCommand cmdTemp = new SqlCommand(strTemp, conn);
            cmdTemp.ExecuteNonQuery();

            SqlBulkCopy sbc = new SqlBulkCopy(conn);
            sbc.DestinationTableName = "[dbo].[##tempBUSA]";
            sbc.ColumnMappings.Add("BUSA", "BUSA");
            sbc.ColumnMappings.Add("Total_Cost", "Total_Cost");
            sbc.ColumnMappings.Add("ODM", "ODM");
            sbc.ColumnMappings.Add("filepath", "filepath");
            sbc.ColumnMappings.Add("UploadTime", "UploadTime");
            sbc.ColumnMappings.Add("version", "version");
            sbc.BulkCopyTimeout = 0;
            sbc.WriteToServer(dtQToolBUSA);
            sbc.Close();

            //**For Quote Tool: BUSA HP Price
            if (DateTime.UtcNow.Day > 12 && DateTime.UtcNow.Day <= 20)  //for final
            {
                string strQPlus_BU = "DELETE FROM [dbo].[QPlus_HP_Price] WHERE [Quote_Type]='Final' and [EffDate]='" + DateTime.UtcNow.ToString("yyyy-MM-01") + "' " +
                    "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct [BUSA], [PMatrix_SA_Desc], 'BUSA', [Total_Cost], [ODM], [PMatrix_Family], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', [filepath], 'BU SA', GETUTCDATE(), 'Eevee', [version], [PMatrix_Category] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [BUSA], [ODM], [PMatrix_Family] ORDER BY [int_BUSA_ID] desc) as [srank] FROM [dbo].[Eevee_BUSA] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "' and [PMatrix_Family] is not null) A " +
                    "WHERE [srank]=1 order by [BUSA], [ODM]";
                SqlCommand cmdQPlus_BU = new SqlCommand(strQPlus_BU, conn);
                cmdQPlus_BU.CommandTimeout = 0;
                cmdQPlus_BU.ExecuteNonQuery();

                string strQPlus_OP = "INSERT INTO [dbo].[QPlus_HP_Price]([DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Config]) " +
                    "SELECT Distinct 'OP', [Cost], [ODM], [PMatrix], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', [filepath], [filesheet], GETUTCDATE(), 'Eevee', [version], [Config_Num] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [ODM], [PMatrix], [Config_Num] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "' and [SA_PN] is not null) A " +
                    "WHERE [srank]=1 order by [filepath], [Config_Num]";
                SqlCommand cmdQPlus_OP = new SqlCommand(strQPlus_OP, conn);
                cmdQPlus_OP.CommandTimeout = 0;
                cmdQPlus_OP.ExecuteNonQuery();

                // + program matrix category???
                string strQPlus_Adder = "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct A.[SA_PN], A.[SA_Desc], 'OP', [TTL], A.[ODM], A.[PMatrix], 'Final', '" + DateTime.UtcNow.ToString("yyyy-MM-01") + "', A.[filepath], A.[filesheet], GETUTCDATE(), 'Eevee', A.[version], A.[SA_Cate] FROM [dbo].[Eevee_Option_NB] A " +
                    "INNER JOIN (SELECT Distinct [SA_PN], [ODM], SUM(t1.[Cost]) as [TTL], [filepath], [version] FROM (SELECT * FROM (SELECT Distinct [SA_PN], [Cost], [ODM], [Config_Option], [filepath], [version], ROW_NUMBER() OVER (PARTITION BY [SA_PN], [ODM], [Config_Option] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] " +
                    "WHERE [SA_PN] is not NULL and [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-12") + "') tb WHERE [srank]=1) as t1 " +
                    "GROUP BY [SA_PN], [ODM], [filepath], [version]) B on B.[SA_PN]=A.[SA_PN] and B.[ODM]=A.[ODM] and B.[filepath]=A.[filepath] and B.[version]=A.[version] order by [PMatrix], [SA_PN], [ODM]";
                SqlCommand cmdQPlus_Adder = new SqlCommand(strQPlus_Adder, conn);
                cmdQPlus_Adder.CommandTimeout = 0;
                cmdQPlus_Adder.ExecuteNonQuery();
            }
            else if (DateTime.UtcNow.Day > 20 && DateTime.UtcNow.Day <= 30)  //for initial
            {
                string strQPlus_BU = "DELETE FROM [dbo].[QPlus_HP_Price] WHERE [Quote_Type]='Initial' and [EffDate]='" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "' " +
                    "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct [BUSA], [PMatrix_SA_Desc], 'BUSA', [Total_Cost], [ODM], [PMatrix_Family], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', [filepath], 'BU SA', GETUTCDATE(), 'Eevee', [version], [PMatrix_Category] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [BUSA], [ODM], [PMatrix_Family] ORDER BY [int_BUSA_ID] desc) as [srank] FROM [dbo].[Eevee_BUSA] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "' and [PMatrix_Family] is not null) A " +
                    "WHERE [srank]=1 order by [BUSA], [ODM]";
                SqlCommand cmdQPlus_BU = new SqlCommand(strQPlus_BU, conn);
                cmdQPlus_BU.CommandTimeout = 0;
                cmdQPlus_BU.ExecuteNonQuery();

                string strQPlus_OP = "INSERT INTO [dbo].[QPlus_HP_Price]([DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Config]) " +
                    "SELECT Distinct 'OP', [Cost], [ODM], [PMatrix], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', [filepath], [filesheet], GETUTCDATE(), 'Eevee', [version], [Config_Num] FROM (" +
                    "SELECT *, ROW_NUMBER() OVER (PARTITION BY [ODM], [PMatrix], [Config_Num] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] WHERE [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "' and [SA_PN] is not null) A " +
                    "WHERE [srank]=1 order by [filepath], [Config_Num]";
                SqlCommand cmdQPlus_OP = new SqlCommand(strQPlus_OP, conn);
                cmdQPlus_OP.CommandTimeout = 0;
                cmdQPlus_OP.ExecuteNonQuery();

                // + program matrix category???
                string strQPlus_Adder = "INSERT INTO [dbo].[QPlus_HP_Price]([PN], [Description], [DataSource], [Price], [ODM], [Platform], [Quote_Type], [EffDate], [filename], [filesheet], [UploadTime], [AddedBy], [revision], [Category]) " +
                    "SELECT Distinct A.[SA_PN], A.[SA_Desc], 'OP', [TTL], A.[ODM], A.[PMatrix], 'Initial', '" + DateTime.UtcNow.AddMonths(1).ToString("yyyy-MM-01") + "', A.[filepath], A.[filesheet], GETUTCDATE(), 'Eevee', A.[version], A.[SA_Cate] FROM [dbo].[Eevee_Option_NB] A " +
                    "INNER JOIN (SELECT Distinct [SA_PN], [ODM], SUM(t1.[Cost]) as [TTL], [filepath], [version] FROM (SELECT * FROM (SELECT Distinct [SA_PN], [Cost], [ODM], [Config_Option], [filepath], [version], ROW_NUMBER() OVER (PARTITION BY [SA_PN], [ODM], [Config_Option] ORDER BY [int_Option_NB_ID] desc) as [srank] FROM [dbo].[Eevee_Option_NB] " +
                    "WHERE [SA_PN] is not NULL and [UploadTime]>'" + DateTime.UtcNow.ToString("yyyy-MM-20") + "') tb WHERE [srank]=1) as t1 " +
                    "GROUP BY [SA_PN], [ODM], [filepath], [version]) B on B.[SA_PN]=A.[SA_PN] and B.[ODM]=A.[ODM] and B.[filepath]=A.[filepath] and B.[version]=A.[version] order by [PMatrix], [SA_PN], [ODM]";
                SqlCommand cmdQPlus_Adder = new SqlCommand(strQPlus_Adder, conn);
                cmdQPlus_Adder.CommandTimeout = 0;
                cmdQPlus_Adder.ExecuteNonQuery();
            }



            ////02: Upload RCTO, RCTO_PN, Adder
            //string strTemp =
            //    "DROP TABLE IF EXISTS ##tempOption " +
            //    "CREATE TABLE ##tempOption([ODM] CHAR(30), [Platform] CHAR(60), [Size] CHAR(10), [Extension] CHAR(20), [Config_Option] CHAR(400), [Config_Rule] CHAR(200), [Config_Num] CHAR(30), [Rule_1] CHAR(60), [Rule_2] CHAR(60), [Rule_3] CHAR(60), " +
            //    "[AV_Category] CHAR(80), [Cost] FLOAT, [rowId] INT, [filepath] CHAR(300), [filesheet] CHAR(50), [UploadTime] SMALLDATETIME, [version] CHAR(20)) " +
            //    "DROP TABLE IF EXISTS ##tempExclude " +
            //    "CREATE TABLE ##tempExclude([Platform] CHAR(60), [Exclude] CHAR(60), [filepath] CHAR(300), [filesheet] CHAR(50))";
            //SqlCommand cmdTemp = new SqlCommand(strTemp, conn);
            //cmdTemp.ExecuteNonQuery();

            //SqlBulkCopy sbcExclude = new SqlBulkCopy(conn);
            //sbcExclude.DestinationTableName = "[dbo].[##tempExclude]";
            //sbcExclude.ColumnMappings.Add("Platform", "Platform");
            //sbcExclude.ColumnMappings.Add("Exclude", "Exclude");
            //sbcExclude.ColumnMappings.Add("filepath", "filepath");
            //sbcExclude.ColumnMappings.Add("filesheet", "filesheet");
            //sbcExclude.BulkCopyTimeout = 0;
            //sbcExclude.WriteToServer(dtExclude);
            //sbcExclude.Close();



            //SqlBulkCopy sbcAdder = new SqlBulkCopy(conn);
            //sbcAdder.DestinationTableName = "[dbo].[##tempOption]";
            //sbcAdder.ColumnMappings.Add("Platform", "Platform");
            //sbcAdder.ColumnMappings.Add("ODM", "ODM");
            //sbcAdder.ColumnMappings.Add("Extension", "Extension");
            //sbcAdder.ColumnMappings.Add("Size", "Size");
            //sbcAdder.ColumnMappings.Add("Config_Option", "Config_Option");
            //sbcAdder.ColumnMappings.Add("Config_Rule", "Config_Rule");
            //sbcAdder.ColumnMappings.Add("Config_Num", "Config_Num");
            //sbcAdder.ColumnMappings.Add("Rule_1", "Rule_1");
            //sbcAdder.ColumnMappings.Add("Rule_2", "Rule_2");
            //sbcAdder.ColumnMappings.Add("Rule_3", "Rule_3");
            //sbcAdder.ColumnMappings.Add("AV_Category", "AV_Category");
            //sbcAdder.ColumnMappings.Add("Cost", "Cost");
            //sbcAdder.ColumnMappings.Add("rowId", "rowId");
            //sbcAdder.ColumnMappings.Add("filepath", "filepath");
            //sbcAdder.ColumnMappings.Add("filesheet", "filesheet");
            //sbcAdder.ColumnMappings.Add("UploadTime", "UploadTime");
            //sbcAdder.ColumnMappings.Add("version", "version");
            //sbcAdder.BulkCopyTimeout = 0;
            //sbcAdder.WriteToServer(dtAdder);
            //sbcAdder.Close();

            ////03: Update vendor code for RCTO, RCTO_PN
            //string strVendorcode =
            //    "UPDATE t1 SET [Vendor_Code]=t2.[Vendor_Code] FROM [dbo].[Eevee_Option_NB_RCTO] t1, [dbo].[BB8_ODM_Lookup] t2 WHERE t1.[version]='" + ver + "' and t2.[Note]='RCTO' and t2.[BU]='NB' and t1.[RCTO_Site_Code]=t2.[Site_Code] and t1.[ODM] LIKE '%'+ RTRIM(t2.[ODM]) +'%' " +
            //    "UPDATE t1 SET [Vendor_Code]=t2.[Vendor_Code] FROM [dbo].[Eevee_Option_NB_RCTO_PN] t1, [dbo].[BB8_ODM_Lookup] t2 WHERE t1.[version]='" + ver + "' and t2.[Note]='RCTO' and t2.[BU]='NB' and t1.[RCTO_Site_Code]=t2.[Site_Code] and t1.[ODM] LIKE '%'+ RTRIM(t2.[ODM]) +'%'"; //and t2.[Note]='RCTO'
            //SqlCommand cmdVendorcode = new SqlCommand(strVendorcode, conn);
            //cmdVendorcode.CommandTimeout = 0;
            //cmdVendorcode.ExecuteNonQuery();

            ////04: Find Adder PN
            //string strFindSA =
            //    "INSERT INTO [dbo].[Eevee_Option_NB]([ODM], [Platform], [Extension], [Size], [Config_Option], [Config_Rule], [Config_Num], [Rule_1], [Rule_2], [Rule_3], [AV_Category], [Cost], [rowId], [PMatrix], [SA_PN], [SA_Desc], [SA_Cate], [filepath], [filesheet], [UploadTime], [version]) " +
            //    "SELECT Distinct t1.[ODM], t1.[Platform], [Extension], [Size], [Config_Option], [Config_Rule], [Config_Num], [Rule_1], [Rule_2], [Rule_3], [AV_Category], t1.[Cost], [rowId], t2.[Family], t2.[SA_Number], t2.[SA_Description], t2.[Category], [filepath], [filesheet], t1.[UploadTime], [version] FROM [dbo].[##tempOption] t1 " +
            //    "LEFT JOIN [dbo].[PMatrix_AVSA_latest] t2 ON t2.[Family] LIKE '' + RTRIM(t1.[Platform]) + '%' and t2.[Family] LIKE '%'+RTRIM(t1.[Extension])+'%' and RTRIM(t2.[Family]) LIKE '%' + LTRIM(RTRIM(t1.[Size])) + '' " +
            //    "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_1]) + '%' collate Chinese_PRC_CS_AI " +
            //    "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_2]) + '%' collate Chinese_PRC_CS_AI " +
            //    "and t2.[SA_Description] LIKE '%' + RTRIM(t1.[Rule_3]) + '%' collate Chinese_PRC_CS_AI " +
            //    "and t2.[Category] LIKE '%' + RTRIM(t1.[AV_Category]) + '%' and LTRIM(t2.[SA_Description]) NOT LIKE 'GNRC%' and LTRIM(t2.[SA_Description]) NOT LIKE 'SSA%' " +
            //    "" +
            //    "DELETE FROM [dbo].[Eevee_Option_NB] WHERE [version]='" + ver + "' and [Size]<>'' and [Size] NOT LIKE '%W%' and [PMatrix] is not NULL and SUBSTRING([PMatrix], LEN([Platform])+1, LEN([PMatrix])) LIKE '%W%' " +
            //    "DELETE FROM [dbo].[Eevee_Option_NB] FROM [dbo].[Eevee_Option_NB] as A, [dbo].[##tempExclude] as B WHERE A.[version]='" + ver + "' and A.[filepath]=B.[filepath] and A.[filesheet]=B.[filesheet] and A.[PMatrix]=B.[Exclude]";
            //SqlCommand cmdFindSA = new SqlCommand(strFindSA, conn);
            //cmdFindSA.CommandTimeout = 0;
            //cmdFindSA.ExecuteNonQuery();

            ////04: Config for PN
            //string strConfigPN = "UPDATE t1 SET t1.[PMatrix]=t2.[Family], t1.[SA_Desc]=t2.[SA_Description], t1.[SA_Cate]=t2.[Category] FROM [dbo].[Eevee_Option_NB_Config_PN] t1, " +
            //    "(SELECT distinct A.*, B.[Family], B.[SA_Number], B.[SA_Description], B.[Category] FROM [dbo].[Eevee_Option_NB_Config_PN] A " +
            //    "LEFT JOIN [dbo].[PMatrix_AVSA_latest] as B on B.[SA_Number]=A.[SA_PN] and B.[Family] LIKE '%'+RTRIM(A.[Platform])+'%' and B.[Family] LIKE '%'+RTRIM(A.[Size])+'%' and B.[Family] LIKE '%'+RTRIM(A.[Extension])+'%' and B.[Category] LIKE '%'+RTRIM(A.[AV_Category])+'%' " +
            //    "WHERE A.[version]='" + ver + "' and B.[Family] is not null) t2 WHERE t1.[int_Config_for_PN_NB_ID]=t2.[int_Config_for_PN_NB_ID] " +
            //    "" +
            //    "DELETE FROM [dbo].[Eevee_Option_NB_Config_PN] WHERE [version]='" + ver + "' and [PMatrix] is null " +
            //    "DELETE FROM [dbo].[Eevee_Option_NB_Config_PN] WHERE [version]='" + ver + "' and [PMatrix] NOT LIKE '%'+RTRIM([Platform])+'%' and [PMatrix] NOT LIKE '%'+RTRIM([Size])+'%' and [PMatrix] NOT LIKE '%'+RTRIM([Extension])+'%'";
            //SqlCommand cmdConfigPN = new SqlCommand(strConfigPN, conn);
            //cmdConfigPN.CommandTimeout = 0;
            //cmdConfigPN.ExecuteNonQuery();

            //string strUpload =
            //   "INSERT INTO [dbo].[Eevee_Upload_Report]([PN], [Sitecode], [Vendorcode], [Cost], [EffDateFrom], [EffDateTo], [BU], [Type], [UploadTime], [filepath], [version]) " +
            //   "SELECT Distinct [SA_PN], [Site_Code], [Vendor_Code], SUM(t1.[Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
            //   "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], [filepath], [version], [Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
            //   "LEFT JOIN [dbo].[BB8_ODM_Lookup] as B on B.[ODM]=A.[ODM] " +
            //   "WHERE [SA_PN] is not NULL and [version]='" + ver + "' and [Note]='IDS' and [BU]='NB') as t1 GROUP BY [SA_PN], [Site_Code], [Vendor_Code], [filepath], [version] " +
            //   "UNION " +
            //   "SELECT [SA_PN], [RCTO_Site_Code], [Vendor_Code], SUM([Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
            //   "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], A.[filepath], A.[version], [RCTO_Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
            //   "LEFT JOIN [dbo].[Eevee_Option_NB_RCTO] as B on B.[filesheet]=A.[filesheet] and B.[filepath]=A.[filepath] and B.[version]=A.[version] " +
            //   "WHERE [SA_PN] is not NULL and [RCTO_Site_Code] is not NULL and A.[version]='" + ver + "') as t1 GROUP BY [SA_PN], [RCTO_Site_Code], [Vendor_Code], [filepath], [version] " +
            //   "UNION " +
            //   "SELECT [SA_PN], [RCTO_Site_Code], [Vendor_Code], SUM([Cost]), GETUTCDATE(), '9999-12-31', 'NB', 'Option', GETUTCDATE(), [filepath], [version] FROM " +
            //   "(SELECT Distinct [SA_PN], A.[ODM], [Config_Option], [Cost], A.[filepath], A.[version], [RCTO_Site_Code], [Vendor_Code] FROM [dbo].[Eevee_Option_NB] A " +
            //   "LEFT JOIN [dbo].[Eevee_Option_NB_RCTO_PN] as B on B.[filesheet]=A.[filesheet] and B.[filepath]=A.[filepath] and B.[version]=A.[version] and B.[rowId]=A.[rowId] and B.[PMatrix]=A.[PMatrix] " +
            //   "WHERE [SA_PN] is not NULL and [RCTO_Site_Code] is not NULL and A.[version]='" + ver + "') as t1 GROUP BY [SA_PN], [RCTO_Site_Code], [Vendor_Code], [filepath], [version] " +
            //   "Order by [filepath], [SA_PN], [Site_Code]";
            //SqlCommand cmdUpload = new SqlCommand(strUpload, conn);
            //cmdUpload.CommandTimeout = 0;
            //cmdUpload.ExecuteNonQuery();



        }




    }
}
