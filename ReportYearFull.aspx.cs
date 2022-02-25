using System;
using System.Collections.Generic;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

using ASK.Models.ModelLayer;
using ASK.ViewModels;

namespace ASK.Pages
{
    public partial class ReportYearFull : System.Web.UI.Page
    {
        YearWorkerFull yearDataWorker = new YearWorkerFull();
        ViewModelYearFull model = new ViewModelYearFull();
        BoundField newColumn;
        GridView currTable;

        BoundField SetVisibilityNewColumn(BoundField newColumn, bool isSensorUse)
        {
            if (isSensorUse)
            {
                newColumn.Visible = true;
            }
            else
            {
                newColumn.Visible = false;
            }
            return newColumn;
        }
        void AddColumnFullYEarTable()
        {
            currTable = EmissionTable;
            currTable.Columns.Clear();
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "Month",
                    HeaderText = "Месяц"
                };
            currTable.Columns.Add(newColumn);
            //==================================================================            
            newColumn =
                new BoundField()
                {
                    DataField = "CO_cnc",
                    HeaderText = "CO мг/нм3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_CO_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "CO_cnc_alfa2",
                    HeaderText = "CO alfa2 мг/нм3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, false);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "CO2_cnc",
                    HeaderText = "CO2 мг/нм3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_CO2_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NO_cnc",
                    HeaderText = "NO мг/м3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NO_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NO2_cnc",
                    HeaderText = "NO2 мг/м3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NO2_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NO2_cnc_alfa1",
                    HeaderText = "NO2 alfa1 мг/м3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, false);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
               new BoundField()
               {
                   DataField = "NO2_cnc_alfa2",
                   HeaderText = "NO2 alfa2 мг/м3",
                   DataFormatString = "{0:F3}"
               };
            newColumn = SetVisibilityNewColumn(newColumn, false);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NOx_cnc",
                    HeaderText = "NOx мг/нм3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NOx_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "SO2_cnc",
                    HeaderText = "SO2 мг/нм3",
                    DataFormatString = "{0:F3}",
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_SO2_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "Dust_cnc",
                    HeaderText = "Тв.ч. мг/нм3",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_Dust_conc_calculated);
            currTable.Columns.Add(newColumn);
            
            //=================================================================================================================================================            
            newColumn =
                new BoundField()
                {
                    DataField = "CO_emiss",
                    HeaderText = "CO г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_CO_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "CO2_emiss",
                    HeaderText = "CO2 г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_CO2_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NO_emiss",
                    HeaderText = "NO г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NO_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NO2_emiss",
                    HeaderText = "NO2 г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NO2_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "NOx_emiss",
                    HeaderText = "NOx г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_NOx_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "SO2_emiss",
                    HeaderText = "SO2 г/с",
                    DataFormatString = "{0:F3}",
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_SO2_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "Dust_emiss",
                    HeaderText = "Тв.ч. г/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askCalculatedConcEmiss.is_Dust_emiss_calculated);
            currTable.Columns.Add(newColumn);
            
            //==================================================================       
            newColumn =
                new BoundField()
                {
                    DataField = "Press",
                    HeaderText = "Давление ДГ, кПа",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askUseSensors.is_Pressure_truba_sensor_use);
            currTable.Columns.Add(newColumn);
            //==================================================================            
            newColumn =
                new BoundField()
                {
                    DataField = "Temper",
                    HeaderText = "Температура ДГ, °C",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askUseSensors.is_Temperature_truba_sensor_use);
            currTable.Columns.Add(newColumn);
            //==================================================================            
            newColumn =
                new BoundField()
                {
                    DataField = "Speed",

                    //HeaderText = "Скорость м/с",
                    HeaderText = "Скорость ДГ, м3/с",

                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, askUseSensors.is_Speed_truba_sensor_use);
            currTable.Columns.Add(newColumn);
            //==================================================================            
            newColumn =
                new BoundField()
                {
                    DataField = "Flow",
                    HeaderText = "Расход  ДГ нм3/с",
                    DataFormatString = "{0:F3}"
                };
            newColumn = SetVisibilityNewColumn(newColumn, true);
            currTable.Columns.Add(newColumn);
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "O2",
                    HeaderText = "O2 сух. %",
                    DataFormatString = "{0:F3}"
                };
            if (askUseSensors.is_O2_dry_sensor_use)
            {
                newColumn.HeaderText = "O2 сух. %";
            }
            else
            {
                newColumn.HeaderText = "O2 вл. %";
            }
            currTable.Columns.Add(newColumn);
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "H2O",
                    HeaderText = "H2O %",
                    DataFormatString = "{0:F3}"
                };
            if (askUseSensors.is_O2_wet_sensor_use)
            {
                if (askUseSensors.is_O2_dry_sensor_use)
                {
                    newColumn.HeaderText = "O2 вл. %";
                }
                else
                {
                    newColumn.HeaderText = "H2O %";
                }
            }
            currTable.Columns.Add(newColumn);
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "Stops",
                    HeaderText = "Простой",
                    DataFormatString = "{0:F3}",
                    Visible = true
                };
            currTable.Columns.Add(newColumn);
            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "percntStopsMonth",
                    HeaderText = "Простоев (% месяц)",
                    DataFormatString = "{0:F3}",
                    Visible = true
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "percntStopsYear",
                    HeaderText = "Простоев (% год)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "summPercntStops",
                    HeaderText = "Простоев (% с нач. года)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "OffTime",
                    HeaderText = "Отключ. (час.)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "OffTimePercentByMonth",
                    HeaderText = "Отключ. (% месяц)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "OffTimePercentByYear",
                    HeaderText = "Отключ. (% год)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "OffTimePercentFullYear",
                    HeaderText = "Отключ. (% c нач. года)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "DaysInMonth",
                    HeaderText = "Дней в мес.",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "HoursInMonth",
                    HeaderText = "Часов в мес.",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "WorkHours",
                    HeaderText = "Рабоч. часов в мес.",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "PercWorkHoursMonth",
                    HeaderText = "Рабоч. часов (% мес.)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "PercWorkHoursYear",
                    HeaderText = "Рабоч. часов (% год)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

            //==================================================================
            newColumn =
                new BoundField()
                {
                    DataField = "PercWorkHoursYearFull",
                    HeaderText = "Рабоч. часов (% с нач. года)",
                    DataFormatString = "{0:F3}",
                    Visible = false
                };
            currTable.Columns.Add(newColumn);

           

            //==================================================================
        }

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                commentNameReport.Text = 
                    askSiteName.getShortASKname + 
                    askSiteName.getShortSourceName +
                    " Годовой отчёт за ";

                DropDownListYear.Text = Convert.ToString(DateTime.Now.Year);
                BindGridValue(DateTime.Now.Year);    
            }
        } // end Page_Load()

        private void BindGridValue(int inYearParam)
        {
            AddColumnFullYEarTable();

            model =  yearDataWorker.getYearReport(inYearParam);

            EmissionTable.DataSource = model.MonthData;
            EmissionTable.DataBind();

            commentYear.Text = DropDownListYear.Text;

            double stopsPercents = Math.Round(Convert.ToDouble(model.stops) * 100, 3);

            commentStopData.Text = stopsPercents.ToString() + "%";

        } // end BindGridValue()

        protected void EmissionTable_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                if (
                    e.Row.Cells[0].Text == DateTime.Now.Month.ToString() &&
                    DropDownListYear.Text == DateTime.Now.Year.ToString()
                    )
                {
                    e.Row.Cells[0].BackColor = System.Drawing.Color.Yellow;
                    e.Row.Cells[0].ToolTip = "Текущий месяц";
                    //e.Row.Cells[12].BackColor = System.Drawing.Color.Gray;
                }
            }



            if (e.Row.RowIndex == 12)
            {
                for (int j = 0; j < e.Row.Cells.Count; j++)
                {
                    e.Row.Cells[j].Font.Bold = true;
                    e.Row.Cells[j].BackColor = System.Drawing.Color.DarkGray;
                }
            }



                if (e.Row.RowType == DataControlRowType.Footer)
            {
                e.Row.Font.Bold = true;
                e.Row.Cells[0].Text = "Cумма (т.):";
                e.Row.Cells[11].Text = Convert.ToString(Math.Round(model.SumData.CO_emiss, 3));
                e.Row.Cells[12].Text = Convert.ToString(Math.Round(model.SumData.CO2_emiss, 3));
                e.Row.Cells[13].Text = Convert.ToString(Math.Round(model.SumData.NO_emiss, 3));
                e.Row.Cells[14].Text = Convert.ToString(Math.Round(model.SumData.NO2_emiss, 3));
                e.Row.Cells[15].Text = Convert.ToString(Math.Round(model.SumData.NOx_emiss, 3));
                e.Row.Cells[16].Text = Convert.ToString(Math.Round(model.SumData.SO2_emiss, 3));
                e.Row.Cells[17].Text = Convert.ToString(Math.Round(model.SumData.Dust_emiss, 3));

                e.Row.Cells[24].Text = Convert.ToString(Math.Round(model.SumData.Stops, 3));
            }

         

        } // end EmissionTable_RowDataBound()

        protected void BtnToExcel_Click(object sender, EventArgs e)
        {
            string NameFile = "attachment; filename=" + DropDownListYear.Text + ".xls";

            Response.ClearContent();
            Response.AddHeader("content-disposition", NameFile);
            Response.ContentType = "application/excel";
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1251");
            //Response.ContentEncoding = Encoding.UTF8;

            System.IO.StringWriter sw = new System.IO.StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            fullComment.RenderControl(htw);
            EmissionTable.RenderControl(htw);

            commentStops.RenderControl(htw);

            Response.Write(sw.ToString());
            Response.End();
        } // end BtnToExcel_Click()
        protected void BtnToExcel_Click2(object sender, EventArgs e)
        {
            Excel.ExcelWorker ExcelWorker = new Excel.ExcelWorker();

            int year = Convert.ToInt32(DropDownListYear.Text);            
            DateTime date = new DateTime(year, 1, 1);
            
            string path = askSiteName.GetFullYearReportPath(date);

            ExcelWorker.Delete(path);
            ExcelWorker.CreateReportYearFull(date, path);
            //=====================================
            //To Get the physical Path of the file

            //string filepath = Server.MapPath("test.txt");
            string filepath = path;

            // Create New instance of FileInfo class to get the properties of the file being downloaded
            System.IO.FileInfo myfile = new System.IO.FileInfo(filepath);

            // Checking if file exists
            if (myfile.Exists)
            {
                // Clear the content of the response
                Response.ClearContent();

                // Add the file name and attachment, which will force the open/cancel/save dialog box to show, to the header
                Response.AddHeader("Content-Disposition", "attachment; filename=" + myfile.Name);

                // Add the file size into the response header
                Response.AddHeader("Content-Length", myfile.Length.ToString());

                // Set the ContentType
                Response.ContentType = ReturnExtension(myfile.Extension.ToLower());

                // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
                Response.TransmitFile(myfile.FullName);

                // End the response
                Response.End();
            }
            //=====================================
        } // end BtnToExcel_Click()

        private string ReturnExtension(string fileExtension)
        {
            switch (fileExtension)
            {
                case ".htm":
                case ".html":
                case ".log":
                    return "text/HTML";
                case ".txt":
                    return "text/plain";
                case ".doc":
                    return "application/ms-word";
                case ".tiff":
                case ".tif":
                    return "image/tiff";
                case ".asf":
                    return "video/x-ms-asf";
                case ".avi":
                    return "video/avi";
                case ".zip":
                    return "application/zip";
                case ".xls":
                case ".xlsx":
                case ".csv":
                    return "application/vnd.ms-excel";
                case ".gif":
                    return "image/gif";
                case ".jpg":
                case "jpeg":
                    return "image/jpeg";
                case ".bmp":
                    return "image/bmp";
                case ".wav":
                    return "audio/wav";
                case ".mp3":
                    return "audio/mpeg3";
                case ".mpg":
                case "mpeg":
                    return "video/mpeg";
                case ".rtf":
                    return "application/rtf";
                case ".asp":
                    return "text/asp";
                case ".pdf":
                    return "application/pdf";
                case ".fdf":
                    return "application/vnd.fdf";
                case ".ppt":
                    return "application/mspowerpoint";
                case ".dwg":
                    return "image/vnd.dwg";
                case ".msg":
                    return "application/msoutlook";
                case ".xml":
                case ".sdxl":
                    return "application/xml";
                case ".xdp":
                    return "application/vnd.adobe.xdp+xml";
                default:
                    return "application/octet-stream";
            }
        }
        protected void ButtonCreateMonthReport_Click(object sender, EventArgs e)
        {
            BindGridValue(Convert.ToInt32(Convert.ToInt32(DropDownListYear.Text)));
        }

        //========================================================================================
        // Пустые события необходимы для корректной работы некоторых элементов и экспорта в Excel

        protected void EmissionTable_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
        }

    } // end class

} // end namespace