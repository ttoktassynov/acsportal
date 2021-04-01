using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.DirectoryServices;
using System.Web.UI.HtmlControls;
using System.Xml;
using System.Globalization;
using System.Security.Principal;

public partial class Monitoring : System.Web.UI.Page
{
    private Excel.Application excelapp;
    private Excel.Workbooks excelworkbooks;
    private Excel.Workbook excelworkbook;
    private Excel.Worksheet excelworksheet1;
    private Excel.Range excelcells10;
    protected void Page_Load(object sender, EventArgs e)
    {
        
        //if (!User.Identity.IsAuthenticated)
        //{
        //    var returnUrl = Server.UrlEncode(Request.Url.PathAndQuery);
        //    Response.Redirect("~/Account/Login.aspx?ReturnURL=" + returnUrl);
        //}

        if (!this.IsPostBack)
        {
           
          
        }
    }


    #region help objects
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    private void DownloadReport(string path)
    {
        FileInfo fileInfo = new FileInfo(path);

        if (fileInfo.Exists)
        {
            try
            {
                Response.Clear();

                Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlPathEncode(fileInfo.Name));
                Response.AddHeader("Content-Length", fileInfo.Length.ToString());
                Response.ContentType = "application/octet-stream";
                //Response.ContentType = "application/vnd.ms-excel";
                Response.Flush();
                Response.WriteFile(fileInfo.FullName);
                Response.End();
            }
            catch (ThreadAbortException exx)
            {

            }

        }
        else
        {

        }
    }
    private void OpenExcelFile(string filename)
    {
        excelapp = new Excel.Application();
        excelapp.Visible = true;
        excelworkbooks = excelapp.Workbooks;
        excelapp.Workbooks.Open(filename, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        excelworkbook = excelworkbooks[1];
    }
    private void SaveToPath(string path)
    {
        excelworkbook = excelworkbooks[1];
        excelapp.DefaultSaveFormat = Excel.XlFileFormat.xlHtml;
        excelapp.DisplayAlerts = false;
        excelworkbook.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    }
   
    private void KillExcelApp()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();

        if (excelapp != null)
        {
            //excelapp.DisplayAlerts = false;
            excelapp.Quit();
            int hWnd = excelapp.Application.Hwnd;
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            Process[] procs = Process.GetProcessesByName("EXCEL");
            foreach (Process p in procs)
            {
                if (p.Id == processID)
                    p.Kill();
            }
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelapp);
            excelapp = null;
        }
    }
    #endregion
    #region nkt reports
    protected void lb_nktreports_Click(object sender, EventArgs e)
    {
        //ContentPlaceHolder cph = (ContentPlaceHolder)this.Master.FindControl("MainContent");
        LinkButton lb = (LinkButton)sender;
        par_reportname_nkt.InnerText = lb.Text;
        par_error_nkt.InnerText = "";
        switch (lb.ID)
        {
            case "lb_narabotka_skv":
                tbl_portyanka_bri.Visible = false;
                tbl_portyanka_skv.Visible = false;
                tbl_narabotka_skv.Visible = true;
                FillDD_CDNG();
                break;
            case "lb_portyanka_skv":
                tbl_portyanka_bri.Visible = false;
                tbl_narabotka_skv.Visible = false;
                tbl_portyanka_skv.Visible = true;
                FillDD_month_portyanka_skv();
                FillDD_CDNG_potyanka_skv();
                break;
            case "lb_portyanka_bri":
                tbl_portyanka_bri.Visible = true;
                tbl_portyanka_skv.Visible = false;
                tbl_narabotka_skv.Visible = false;
                FillDD_month();
                break;
            case "lb_zameri_skv":
                tbl_portyanka_bri.Visible = false;
                tbl_narabotka_skv.Visible = false;
                tbl_portyanka_skv.Visible = false;
                break;

        }
    }
    protected void FillDD_month()
    {
        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        DataTable dt = db.GetDataTable("SELECT monthid,monthnamerus FROM MONTHS ORDER BY monthid");
        ddl_nktreportmonth.DataSource = dt;
        ddl_nktreportmonth.DataTextField = "monthnamerus";
        ddl_nktreportmonth.DataValueField = "monthid";
        ddl_nktreportmonth.DataBind();
    }
    protected void FillDD_month_portyanka_skv()
    {
        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        DataTable dt = db.GetDataTable("SELECT monthid,monthnamerus FROM MONTHS ORDER BY monthid");
        ddl_month_portyanka_skv.DataSource = dt;
        ddl_month_portyanka_skv.DataTextField = "monthnamerus";
        ddl_month_portyanka_skv.DataValueField = "monthid";
        ddl_month_portyanka_skv.DataBind();
    }
    protected void FillDD_CDNG()
    {
        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        DataTable dt = db.GetDataTable("SELECT cdngno,cdngname FROM CDNG where ngduid = " + ddl_ngdu_nar.SelectedItem.Value.ToString() + " ORDER BY cdngno");
        ddl_cdng_nar.DataSource = dt;
        ddl_cdng_nar.DataTextField = "cdngname";
        ddl_cdng_nar.DataValueField = "cdngno";
        ddl_cdng_nar.DataBind();
    }
    protected void FillDD_CDNG_potyanka_skv()
    {
        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        DataTable dt = db.GetDataTable("SELECT cdngno,cdngname FROM CDNG where ngduid = " + ddl_ngdu_portyanka_skv.SelectedItem.Value.ToString() + " ORDER BY cdngno");
        ddl_cdng_portyanka_skv.DataSource = dt;
        ddl_cdng_portyanka_skv.DataTextField = "cdngname";
        ddl_cdng_portyanka_skv.DataValueField = "cdngno";
        ddl_cdng_portyanka_skv.DataBind();
    }
    protected void ddl_ngdu_nas_SelectedIndexChanged(object sender, EventArgs e)
    {
        FillDD_CDNG();
    }
    protected void ddl_ngdu_portyanka_skv_OnSelectedIndexChanged(object sender, EventArgs e)
    {
        FillDD_CDNG_potyanka_skv();
    }
    protected void btn_getportyankareport_Click(object sender, EventArgs e)
    {
        try
        {

            string ngduname = ddl_ngdu.SelectedItem.Text;

            string filename = MapPath(@"~\Шаблоны\Сводка по проведенным ПРС по бригадам " + ngduname + ".xlsx");
            OpenExcelFile(filename);

            excelworksheet1 = (Excel.Worksheet)excelworkbook.Sheets[1];
            
            ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));

            int i =1;
            int j = 1;
            int addpos = 0;
            int currow = 0;
            string curbrig = "";
            int curday = 0;
            int curmonth = 0;
            int curyear = 0;
            DateTime curdate;
            DataTable dt;

            curmonth = Int32.Parse(ddl_nktreportmonth.SelectedItem.Value);
            curyear = Int32.Parse(ddl_nktreportyear.SelectedItem.Value);
            for (i = 1; i <=15;i ++)
            {
                if (i == 1)
                {
                    addpos = 0;
                }
                if (i == 6)
                {
                    addpos = 2;
                }
                if (i == 11)
                {
                    addpos = 4;
                }
                currow = 5 + addpos + (i - 1) * 4;
                curbrig = ((Excel.Range)excelworksheet1.Cells[currow, 3]).Text.ToString();
                for (j = 1; j <= 30; j++)
                {
                    curday = Int32.Parse(((Excel.Range)excelworksheet1.Cells[4, j + 4]).Text.ToString());
                    curdate = new DateTime(curyear, curmonth, curday);
                    dt = db.GetDataTable("SELECT  Скважина, ГУ, [Сост НКТ],ДлинаМ, МРП  FROM CITS where Бригада  = '" + curbrig + "' and Day(Дата)=" + curday + " and Year(Дата) =" + curyear + " and Month(Дата)=" + curmonth);
                    if (dt.Rows.Count != 0)
                    {
                        ((Excel.Range)excelworksheet1.Cells[currow + 0, j + 4]).Value = dt.Rows[0].ItemArray[1] + "/" + dt.Rows[0].ItemArray[0];
                        ((Excel.Range)excelworksheet1.Cells[currow + 0, j + 4]).Font.Bold = true;
                        ((Excel.Range)excelworksheet1.Cells[currow + 1, j + 4]).Value = "Наработка: " + dt.Rows[0].ItemArray[2] + " - " + dt.Rows[0].ItemArray[4];
                        ((Excel.Range)excelworksheet1.Cells[currow + 2, j + 4]).Value = "Обновление: " + dt.Rows[0].ItemArray[3];
                        ((Excel.Range)excelworksheet1.Cells[currow + 3, j + 4]).Value = "Ремонт себебі: ";
                    }

                }
            }


            string path = MapPath("~/Reports/Сводка по проведенным ПРС по бригадам " + ngduname + " за " + ddl_nktreportmonth.SelectedItem.Text + " "+ddl_nktreportyear.SelectedItem.Text + " год.xlsx");
            string file = Path.GetFileName(path);

            SaveToPath(path);
            KillExcelApp();
            DownloadReport(path);
            par_error_nkt.InnerText = "Отчет готов. ОН находится в Мои документы ->  Загрузки";

        }
        catch (Exception ex)
        {
            par_error_nkt.InnerText = "Попробуйте еще раз. Тескт ошибки: " + ex.Message;
            KillExcelApp();
        }
    }
    protected void btn_get_narabotka_skv_Click(object sender, EventArgs e)
    {
        string cdng = ddl_cdng_nar.SelectedItem.Text;
        string fond = ddl_fond_nar.SelectedItem.Value.ToString(); ;

        string filename = MapPath(@"~\Шаблоны\Наработка по скважинам.xlsx");
        OpenExcelFile(filename);
        excelworksheet1 = (Excel.Worksheet)excelworkbook.Sheets[1];
        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        int incol = 8;

        int curskv;
        int prevskv;
        DateTime curdate;
        int curmonthint;
        int curdayint;
        int curyearint;
        int reccount;

        double curdlina;
        string curgu;
        string curmrp;
        int counter;
        string curvid;
        string curtech;
        string cursost;

        counter = 0;
        int narabotka;
        int shift;

        DataTable rsSkv = db.GetDataTable("SELECT distinct Скважина  FROM CITS where ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') order by Скважина ");
        if (rsSkv.Rows.Count != 0)
        {
            reccount = rsSkv.Rows.Count;
        }
        else
        {
            reccount = 0;
        }
        int i;
        i = 84;
        prevskv = 0;
        narabotka = 0;
        shift = 0;
        ((Excel.Range)excelworksheet1.Cells[44, 3]).Value = cdng;
        int k;
        int cmonth=0;
        int cyear=0;
        DataTable rskol;
        #region top report
        for (k = 1; k <= 16; k++)
        {
            if (k >= 1 & k <= 4)
            {
                cyear = 2012;
                cmonth = 8 + k;
            }
            if (k > 4)
            {
                cyear = 2013;
                cmonth = k - 4;
            }
            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[4, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[21, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[21, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and ЦДНГ = '" + cdng + "' and [Сост НКТ] ='б/у' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[5, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Сост НКТ] ='б/у' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[22, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[22, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Сост НКТ] ='новые'  and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[6, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Сост НКТ] ='новые'  and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[23, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[23, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            //shgn kolvo
            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[7, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[24, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[24, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='б/у' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[8, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='б/у' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[25, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[25, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and  [Сост НКТ] ='новые' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[9, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT SUM(МРП)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and  [Сост НКТ] ='новые' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
            if (rskol.Rows.Count != 0)
            {
                if (rskol.Rows[0].ItemArray[0].ToString() != "")
                {
                    ((Excel.Range)excelworksheet1.Cells[26, incol + k]).Value = rskol.Rows[0].ItemArray[0];
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[26, incol + k]).Value = 0;
                }
            }
            rskol.Clear();

            //123123123123123
            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[56, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='б/у' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[57, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='новые' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'УЭЦН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[58, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + "  and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[59, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='б/у' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[60, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();

            rskol = db.GetDataTable("SELECT count(*)   FROM CITS where Год = " + cyear + " and Месяц = " + cmonth + " and [Сост НКТ] ='новые' and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "' and [Технология добычи] = 'ШГН' and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.') ");
            if (rskol.Rows.Count != 0)
            {
                ((Excel.Range)excelworksheet1.Cells[61, incol + k]).Value = rskol.Rows[0].ItemArray[0];
            }
            rskol.Clear();


        }
        #endregion
        int i1;
        int j1;
        
        for (i1 = 1; i1 <= 6; i1++)
        {
            for (j1 = 1; j1 <= 16; j1++)
            {
                Double prnar  = Double.Parse(((Excel.Range)excelworksheet1.Cells[37 + i1, incol + j1]).Value.ToString());
                string prnarstr = prnar.ToString("###0");

                ((Excel.Range)excelworksheet1.Cells[46 + i1, incol + j1]).Value = ((Excel.Range)excelworksheet1.Cells[3 + i1, incol + j1]).Value.ToString() + "/" + ((Excel.Range)excelworksheet1.Cells[55 + i1, incol + j1]).Value.ToString() + "/" + prnarstr;
            }
        }

        int curcol=0;
        DataTable rs2 = db.GetDataTable("SELECT Скважина,Дата,ГУ, МРП,ДлинаМ,[Технология добычи],[Сост НКТ]  FROM CITS where  ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "'  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') order by Скважина ");
        if (rs2.Rows.Count != 0)
        {
            for (int ii = 0; ii < rs2.Rows.Count; ii++)
            {
                if (rs2.Rows[ii].ItemArray[0].ToString() != "")
                {
                    curskv = Int32.Parse(rs2.Rows[ii].ItemArray[0].ToString());
                }
                else
                {
                    continue;
                }
                if (rs2.Rows[ii].ItemArray[1].ToString() != "")
                {
                    curdate = DateTime.Parse(rs2.Rows[ii].ItemArray[1].ToString());
                }
                else
                {
                    continue;
                }
                if (rs2.Rows[ii].ItemArray[5].ToString() != "")
                {
                    curtech = rs2.Rows[ii].ItemArray[5].ToString();
                }
                else
                {
                    curtech = "";
                }
                if (rs2.Rows[ii].ItemArray[6].ToString() != "")
                {
                    cursost = rs2.Rows[ii].ItemArray[6].ToString();
                }
                else
                {
                    cursost = "";
                }

                if (rs2.Rows[ii].ItemArray[2].ToString() != "")
                {
                    curgu = rs2.Rows[ii].ItemArray[2].ToString();
                }
                else
                {
                    curgu = "";
                }
                if (rs2.Rows[ii].ItemArray[3].ToString() != "")
                {
                    curmrp = rs2.Rows[ii].ItemArray[3].ToString() + " д.";
                }
                else
                {
                    curmrp = "";
                }
                if (rs2.Rows[ii].ItemArray[4].ToString() != "")
                {
                    curdlina = Double.Parse(rs2.Rows[ii].ItemArray[4].ToString());
                }
                else
                {
                    curdlina = 0;
                }
                //if (rs2.Rows[i].ItemArray[5].ToString() != "")
                //{
                  //  curvid = rs2.Rows[i].ItemArray[5].ToString();
                //}
                //else
                //{
                    //curvid = "";
                //}
                curmonthint = curdate.Month;
                curdayint = curdate.Day;
                curyearint = curdate.Year;
                if (curyearint == 2012)
                {
                    curcol = curmonthint;
                }
                if (curyearint ==2013)
                {
                    curcol = 12 + curmonthint;
                }
                if (prevskv != curskv)
                {
                    i = i + 1;
                    counter = counter + 1;
                }
                if (counter < (int)(reccount / 2 + 1))
                {
                    shift = 0;
                }
                if (counter == (int)(reccount / 2 + 1))
                {
                    shift = 0;
                }
                if (cursost == "новые")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Font.Color = System.Drawing.Color.Red;
                }
                rskol = db.GetDataTable("SELECT AVG(МРП)   FROM CITS where ЦДНГ = '" + cdng + "' and [Сост НКТ] = 'б/у' and [Скважина] =  " + curskv + "  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
                if (rskol.Rows.Count != 0)
                {
                    if (rskol.Rows[0].ItemArray[0].ToString() != "")
                    {
                        Double val1 = Double.Parse(rskol.Rows[0].ItemArray[0].ToString());
                        ((Excel.Range)excelworksheet1.Cells[i, incol - 1]).Value = val1.ToString("####.0");
                    }
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, incol - 1]).Value = 0;
                }
                rskol.Clear();
                rskol = db.GetDataTable("SELECT AVG(МРП)   FROM CITS where ЦДНГ = '" + cdng + "' and [Сост НКТ] = 'новые' and [Скважина] =  " + curskv + "  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
                if (rskol.Rows.Count != 0)
                {
                    if (rskol.Rows[0].ItemArray[0].ToString() != "")
                    {
                        Double val2 = Double.Parse(rskol.Rows[0].ItemArray[0].ToString());
                        ((Excel.Range)excelworksheet1.Cells[i, incol]).Value = val2.ToString("####.0");
                    }
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, incol ]).Value = 0;
                }
                rskol.Clear();

                if (((Excel.Range)excelworksheet1.Cells[i, curcol]).Text == "")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Value = ((Excel.Range)excelworksheet1.Cells[i, curcol]).Text + curdate.ToString("dd.MM") + "-" + curmrp;
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Value = ((Excel.Range)excelworksheet1.Cells[i, curcol]).Text + "\n" + curdate.ToString("dd.MM") + "-" + curmrp;
                }
                if (curmrp == "")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Interior.ColorIndex = 43;
                }
                ((Excel.Range)excelworksheet1.Cells[i, 1]).Value = counter;
                ((Excel.Range)excelworksheet1.Cells[i, 2]).Value = "UZN_" + curskv.ToString("000#");
                ((Excel.Range)excelworksheet1.Cells[i, 3]).Value = curgu;

                DataTable rsTR = db.GetDataTable("SELECT [Qж, м3],[Qн, т/сут]  FROM TREJIME where Скв = " + curskv + " and ЦДНГ = '" + cdng + "'");
                if (rsTR.Rows.Count != 0)
                {
                    ((Excel.Range)excelworksheet1.Cells[i, 4]).Value = rsTR.Rows[0].ItemArray[0].ToString();
                    ((Excel.Range)excelworksheet1.Cells[i, 5]).Value = rsTR.Rows[0].ItemArray[1].ToString();
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, 4]).Value = "-";
                    ((Excel.Range)excelworksheet1.Cells[i, 5]).Value = "-";
                }
                rsTR.Clear();
                ((Excel.Range)excelworksheet1.Cells[i, 6]).Value = curtech;

                ((Excel.Range)excelworksheet1.Range[excelworksheet1.Cells[i, 1], excelworksheet1.Cells[i, 24]]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                prevskv = curskv;
            }
        }
        string path = MapPath("~/Reports/Наработка по скважинам по  " +ddl_ngdu_nar.SelectedItem.Text  + "  " + ddl_cdng_nar.SelectedItem.Text+ " " + ddl_year_nar.SelectedItem.Text + " год.xlsx");
        string file = Path.GetFileName(path);

        SaveToPath(path);
        KillExcelApp();
        DownloadReport(path);
        par_error_nkt.InnerText = "Отчет готов. ОН находится в Мои документы ->  Загрузки";

    }
    protected void btn_getportyankaskvreport_Click(object sender, EventArgs e)
    {
        string cdng = ddl_cdng_portyanka_skv.SelectedItem.Text;
        string fond = ddl_fond_portyanka_skv.SelectedItem.Value.ToString();

        string filename = MapPath(@"~\Шаблоны\Сводка по проведенным прс По скважинам.xlsx");
        OpenExcelFile(filename);

        excelworksheet1 = (Excel.Worksheet)excelworkbook.Sheets[1];

        ACSDB db = new ACSDB((System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString));
        DataTable rs2;
        DataTable rsSkv;
        DataTable rskol;
        int monthint = Int32.Parse(ddl_month_portyanka_skv.SelectedItem.Value.ToString());
        string monthstr = ddl_month_portyanka_skv.SelectedItem.Text;
        int incol = 8;

        rs2 = db.GetDataTable("SELECT Скважина,Дата,ГУ, МРП,ДлинаМ,[Технология добычи],[Сост НКТ],Бригада  FROM CITS where Month(Дата) = "+monthint+" and ЦДНГ = '" + cdng + "' and [Тип фонда] = '" + fond + "'  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') order by Скважина ");
        int curskv;
        int prevskv;
        DateTime curdate;
        int curdayint;
        double curdlina;
        string curgu;
        string curmrp;
        string curbrig;
        int counter;
        string curvid;
        string curtech;
        string cursost;

        counter = 0;
        int narabotka;
        int i;
        i = 13;
        prevskv = 0;
        narabotka = 0;
        ((Excel.Range)excelworksheet1.Cells[3, 3]).Value = cdng;
        ((Excel.Range)excelworksheet1.Cells[4, 3]).Value = fond;
        ((Excel.Range)excelworksheet1.Cells[5, 3]).Value = monthstr;
        int k;
        int curcol;
        if (rs2.Rows.Count != 0)
        {
            for (int ii = 0; ii < rs2.Rows.Count; ii++)
            {
                if (rs2.Rows[ii].ItemArray[0].ToString() != "")
                {
                    curskv = Int32.Parse(rs2.Rows[ii].ItemArray[0].ToString());
                }
                else
                {
                    continue;
                }
                if (rs2.Rows[ii].ItemArray[1].ToString() != "")
                {
                    curdate = DateTime.Parse(rs2.Rows[ii].ItemArray[1].ToString());
                }
                else
                {
                    continue;
                }
                if (rs2.Rows[ii].ItemArray[5].ToString() != "")
                {
                    curtech = rs2.Rows[ii].ItemArray[5].ToString();
                }
                else
                {
                    curtech = "";
                }
                if (rs2.Rows[ii].ItemArray[6].ToString() != "")
                {
                    cursost = rs2.Rows[ii].ItemArray[6].ToString();
                }
                else
                {
                    cursost = "";
                }

                if (rs2.Rows[ii].ItemArray[2].ToString() != "")
                {
                    curgu = rs2.Rows[ii].ItemArray[2].ToString();
                }
                else
                {
                    curgu = "";
                }
                if (rs2.Rows[ii].ItemArray[3].ToString() != "")
                {
                    curmrp = rs2.Rows[ii].ItemArray[3].ToString() + " д.";
                }
                else
                {
                    curmrp = "";
                }
                if (rs2.Rows[ii].ItemArray[4].ToString() != "")
                {
                    curdlina = Double.Parse(rs2.Rows[ii].ItemArray[4].ToString());
                }
                else
                {
                    curdlina = 0;
                }
                if (rs2.Rows[ii].ItemArray[7].ToString() != "")
                {
                    curbrig = rs2.Rows[ii].ItemArray[7].ToString();
                }
                else
                {
                    curbrig = "";
                }
                //if (rs2.Rows[i].ItemArray[5].ToString() != "")
                //{
                //  curvid = rs2.Rows[i].ItemArray[5].ToString();
                //}
                //else
                //{
                //curvid = "";
                //}
                curdayint = curdate.Day;
                curcol = 8 + curdayint;
                if (prevskv!= curskv){
                    i = i + 1;
                    counter = counter + 1;
                }
                if (cursost == "новые")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Font.Color = System.Drawing.Color.Red;
                }
                rskol = db.GetDataTable("SELECT AVG(МРП)   FROM CITS where Month(Дата) = "+monthint+" and ЦДНГ = '" + cdng + "' and [Сост НКТ] = 'б/у' and [Скважина] =  " + curskv + "  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
                if (rskol.Rows.Count != 0)
                {
                    if (rskol.Rows[0].ItemArray[0].ToString() != "")
                    {
                        Double val1 = Double.Parse(rskol.Rows[0].ItemArray[0].ToString());
                        ((Excel.Range)excelworksheet1.Cells[i, incol - 1]).Value = val1.ToString("####.0");
                    }
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, incol - 1]).Value = 0;
                }
                rskol.Clear();
                rskol = db.GetDataTable("SELECT AVG(МРП)   FROM CITS where Month(Дата) ="+ monthint+" and  ЦДНГ = '" + cdng + "' and [Сост НКТ] = 'новые' and [Скважина] =  " + curskv + "  and Период in ('< 30 дн.','> 200 дн.','101-200дн.','31-100 дн.','ПР не было') ");
                if (rskol.Rows.Count != 0)
                {
                    if (rskol.Rows[0].ItemArray[0].ToString() != "")
                    {
                        Double val2 = Double.Parse(rskol.Rows[0].ItemArray[0].ToString());
                        ((Excel.Range)excelworksheet1.Cells[i, incol]).Value = val2.ToString("####.0");
                    }
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, incol]).Value = 0;
                }
                rskol.Clear();

                if (((Excel.Range)excelworksheet1.Cells[i, curcol]).Text == "")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Value = ((Excel.Range)excelworksheet1.Cells[i, curcol]).Text + curdate.ToString("dd.MM") + "-" + curmrp+ "Бр:"+curbrig;
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Value = ((Excel.Range)excelworksheet1.Cells[i, curcol]).Text + "\n" + curdate.ToString("dd.MM") + "-" + curmrp + "Бр:" + curbrig;
                }
                if (curmrp == "")
                {
                    ((Excel.Range)excelworksheet1.Cells[i, curcol]).Interior.ColorIndex = 43;
                }
                ((Excel.Range)excelworksheet1.Cells[i, 1]).Value = counter;
                ((Excel.Range)excelworksheet1.Cells[i, 2]).Value = "UZN_" + curskv.ToString("000#");
                ((Excel.Range)excelworksheet1.Cells[i, 3]).Value = curgu;

                DataTable rsTR = db.GetDataTable("SELECT [Qж, м3],[Qн, т/сут]  FROM TREJIME where Скв = " + curskv + " and ЦДНГ = '" + cdng + "'");
                if (rsTR.Rows.Count != 0)
                {
                    ((Excel.Range)excelworksheet1.Cells[i, 4]).Value = rsTR.Rows[0].ItemArray[0].ToString();
                    ((Excel.Range)excelworksheet1.Cells[i, 5]).Value = rsTR.Rows[0].ItemArray[1].ToString();
                }
                else
                {
                    ((Excel.Range)excelworksheet1.Cells[i, 4]).Value = "-";
                    ((Excel.Range)excelworksheet1.Cells[i, 5]).Value = "-";
                }
                rsTR.Clear();
                ((Excel.Range)excelworksheet1.Cells[i, 6]).Value = curtech;

                ((Excel.Range)excelworksheet1.Range[excelworksheet1.Cells[i, 1], excelworksheet1.Cells[i, 39]]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                prevskv = curskv;
            }
        }
        string path = MapPath("~/Reports/Сводка по проведенным ПРС по скважинам по " + ddl_cdng_portyanka_skv.SelectedItem.Text + "  " + ddl_year_portyanka_skv.SelectedItem.Text + " год.xlsx");
        string file = Path.GetFileName(path);

        SaveToPath(path);
        KillExcelApp();
        DownloadReport(path);
        par_error_nkt.InnerText = "Отчет готов. ОН находится в Мои документы ->  Загрузки";


    }

    #endregion 
}