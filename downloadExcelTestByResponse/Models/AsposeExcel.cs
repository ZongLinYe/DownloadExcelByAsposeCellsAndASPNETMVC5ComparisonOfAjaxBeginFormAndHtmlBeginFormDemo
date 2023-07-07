using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using Dapper;
using System.Data.SqlClient;
using System.Configuration;
using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;
using System.IO;

namespace downloadExcelByAsposeComparisonOfAjaxBeginFormAndHtmlBeginForm.Models
{
    public class AsposeExcel
    {


        private Style styleHeader;
        private Style styleBody;
        private Style styleBodyDate;
        private string strSql;
        private string strConnMain;

        public AsposeExcel()
        {
            License license = new License();
            license.SetLicense("Aspose.Total.655.lic");
            InitStyles();
            strSql = "";
            strConnMain = ConfigurationManager.ConnectionStrings["MainDBConnection"].ConnectionString;
        }

        [Required]
        [Display(Name = "帳號")]
        public string USER_ID { get; set; }

        [Required]
        [Display(Name = "姓名")]
        public string USER_NAME { get; set; }

        [Display(Name = "電子郵件")]
        [EmailAddress]
        [Required]
        public string USER_EMAIL { get; set; }

        [Display(Name = "密碼")]
        public string USER_PWD { get; set; }
        
        [Display(Name = "確認密碼")]
        [Compare("USER_PWD")]
        public string CONFIRM_USER_PWD { get; set; }
        public string USER_IP { get; set; }
        public string AGENT_ID { get; set; }

        [Display(Name = "是否刪除")]
        [Required]
        public bool DEL_FLG { get; set; }

        [Display(Name = "更新日期")]
        public DateTime MDF_DATE { get; set; }

        public DateTime CRT_DATE { get; set; }
        public string qType { get; set; }



        internal byte[] DownloadExcel()
        {
            var loggedUser = new
            {
                USER_ID = "SYSTEM",
                USER_NAME= "系統管理員"

            };
            strSql = @"SELECT
	                    user_id,
	                    user_name,
	                    user_email,
	                    del_flg,
	                    crt_date,
	                    mdf_date
                    FROM EMP_USER_COPY";

            using(var cn =new SqlConnection(strConnMain))
            {
                cn.Open();
                var userDatas = cn.Query<AsposeExcel>(strSql).ToList();             


                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells.Merge(0, 0, 1, 6);
                Cell mergedCell = worksheet.Cells[0, 0];

       
                mergedCell.PutValue("測試匯出報表");

                worksheet.Cells["A2"].PutValue("列印人員ID");
                worksheet.Cells["B2"].PutValue(loggedUser.USER_ID);
                worksheet.Cells["C2"].PutValue("列印人員姓名");
                worksheet.Cells["D2"].PutValue(loggedUser.USER_NAME);

                worksheet.Cells["A4"].PutValue("帳號");
                worksheet.Cells["B4"].PutValue("姓名");
                worksheet.Cells["C4"].PutValue("電子信箱");
                worksheet.Cells["D4"].PutValue("是否刪除");
                worksheet.Cells["E4"].PutValue("建立日期");
                worksheet.Cells["F4"].PutValue("更新日期");

                for (int i = 0; i < userDatas.Count; i++)
                {
                    AsposeExcel item = userDatas[i];

                    worksheet.Cells[$"A{i + 5}"].PutValue(item.USER_ID);
                    worksheet.Cells[$"B{i + 5}"].PutValue(item.USER_NAME);
                    worksheet.Cells[$"C{i + 5}"].PutValue(item.USER_EMAIL);
                    worksheet.Cells[$"D{i + 5}"].PutValue(item.DEL_FLG ? "否" : "是");
                    worksheet.Cells[$"E{i + 5}"].PutValue(item.CRT_DATE.ToString("yyyy/MM/dd"));
                    worksheet.Cells[$"F{i + 5}"].PutValue(item.MDF_DATE.ToString("yyyy/MM/dd"));
                }

                var style = workbook.CreateStyle();
                style.Font.Name = "標楷體";

                for (int row = 0; row < worksheet.Cells.Rows.Count; row++)
                {
                    for (int col = 0; col < worksheet.Cells.Columns.Count; col++)
                    {
                        worksheet.Cells[row, col].SetStyle(style);
                    }
                }

                //var tempFile = Path.GetTempFileName();
                //workbook.Save(tempFile, Aspose.Cells.SaveFormat.Xlsx);
                //return tempFile;

                MemoryStream ms = new MemoryStream();
                workbook.Save(ms, Aspose.Cells.SaveFormat.Xlsx);

                return ms.ToArray();
            }           
        }

        private void InitStyles()
        {
            CellsFactory cellsFactory = new CellsFactory();
            styleHeader = cellsFactory.CreateStyle();
            styleHeader.HorizontalAlignment = TextAlignmentType.Center;
            styleHeader.Font.IsBold = true;
            styleHeader.Pattern = BackgroundType.Gray25;
            SetBorder(styleHeader);
            styleBody = cellsFactory.CreateStyle();
            styleBody.Pattern = BackgroundType.Solid;
            SetBorder(styleBody);
            styleBodyDate = cellsFactory.CreateStyle();
            styleBodyDate.Copy(styleBody);
            styleBodyDate.Custom = "yyyy/MM/dd";
        }

        private void SetBorder(Style style)
        {
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
        }


    }
}