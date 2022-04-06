using ExcelDataReader;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ChamCong
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        string PathName = "";
        string PathTofolder = "";

        public DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        private void txb_Path_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txb_Path.Text = openFileDialog.SafeFileName;
                PathName = openFileDialog.FileName;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            List<DataExport> lsExport = new List<DataExport>();

            string NameFile = "Workshifts_" + DateTime.Now.ToLongDateString().Replace("/", "") + DateTime.Now.ToLongTimeString().Replace(":", "").Replace(" ", "") + ".xlsx";
            string SaveFilePath = Path.Combine(PathTofolder, NameFile);

            DataSet ds;

            using (var stream = File.Open(PathName, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);


                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                reader.Close();
            }
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage pck = new ExcelPackage())
            {

                pck.Workbook.Properties.Author = "阮林山";
                pck.Workbook.Properties.Company = "FHS";
                pck.Workbook.Properties.Title = "Exported by 阮林山";
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");
                //Định dạng toàn Sheet
                ws.Cells.Style.Font.Name = "Times New Roman";
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells.Style.Font.Size = 14;

                ws.Column(1).Width = 15;
                ws.Column(2).Width = 15;
                ws.Column(3).Width = 15;
                ws.Column(1).Style.Numberformat.Format = "MM/dd hh:mm";
                ws.Column(2).Style.Numberformat.Format = "MM/dd hh:mm";


                int flag = 7;

                // Kiểm tra từng sheet để tính từng người
                for (int k = 0; k < ds.Tables.Count; k++)
                {

                    if (k == 2)
                    {

                    }

                    DataTable dtData = ds.Tables[k];

                    DataExport dataExport = new DataExport();

                    string Name = dtData.Rows[1][3].ToString().Replace("出勤人員:", "");
                    string CMND = dtData.Rows[0][4].ToString();
                    string UserId = dtData.Rows[3][0].ToString();

                    dataExport.Name = Name;
                    dataExport.CCCD = CMND;
                    dataExport.UserId = UserId;

                    List<DataRaw> lsDataSheet = new List<DataRaw>();

                    int countData = dtData.Rows.Count;
                    for (int j = 3; j < countData; j++)
                    {
                        DataRaw data = new DataRaw();
                        data.DateVao = Convert.ToDateTime(dtData.Rows[j][3]);
                        data.DateVaoThuc = Convert.ToDateTime(dtData.Rows[j][3]);
                        data.DateRa = Convert.ToDateTime(dtData.Rows[j][4]);
                        data.DateRaThuc = Convert.ToDateTime(dtData.Rows[j][4]);
                        data.Time = Convert.ToDouble(dtData.Rows[j][5]);

                        lsDataSheet.Add(data);
                    }

                    // các mốc thời gian để so sánh
                    DateTime Moc730 = new DateTime(2022, 01, 01, 07, 30, 00);
                    DateTime Moc1130 = new DateTime(2022, 01, 01, 11, 30, 00);
                    DateTime Moc1300 = new DateTime(2022, 01, 01, 13, 00, 00);
                    DateTime Moc2330 = new DateTime(2022, 01, 01, 23, 30, 00);
                    DateTime Moc1700 = new DateTime(2022, 01, 01, 17, 00, 00);
                    DateTime Moc1730 = new DateTime(2022, 01, 01, 17, 30, 00);

                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        DateTime dateVao = lsDataSheet[i].DateVao;
                        int monthVao = dateVao.Month;
                        int dayVao = dateVao.Day;
                        int hourVao = dateVao.Hour;
                        int minVao = dateVao.Minute;


                        // Làm tròn thời gian dữ liệu quẹt thẻ_ Vào
                        if (dateVao.TimeOfDay < Moc730.TimeOfDay)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, 7, 30, 00);
                        }
                        else if (minVao > 30)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao + 1, 00, 00);
                        }
                        else if (minVao < 30 && minVao != 0)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao, 30, 00);
                        }
                        else if (minVao == 00)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao, 00, 00);
                        }

                        if (Moc1130.TimeOfDay <= dateVao.TimeOfDay && dateVao.TimeOfDay < Moc1300.TimeOfDay)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, 13, 00, 00);
                        }


                        // Làm tròn thời gian dữ liệu quẹt thẻ_ RA
                        DateTime dateRa = lsDataSheet[i].DateRa;
                        int monthRa = dateRa.Month;
                        int dayRa = dateRa.Day;
                        int hourRa = dateRa.Hour;
                        int minRa = dateRa.Minute;

                        if (dateRa.TimeOfDay > Moc2330.TimeOfDay)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, 23, 30, 00);
                        }
                        else if (minRa < 30 && minRa != 0)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 00, 00);
                        }
                        else if (minRa > 30)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 30, 00);
                        }
                        else if (minRa == 00)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 00, 00);
                        }

                        if (Moc1130.TimeOfDay < dateRa.TimeOfDay && dateRa.TimeOfDay <= Moc1300.TimeOfDay)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, 11, 30, 00);
                        }

                    }

                    // Lấy các dữ liệu bất thường 
                    var lsBatThuong = lsDataSheet.Where(r => r.DateRa.Day != r.DateVao.Day).Select(r => new { r.DateVaoThuc, r.DateRaThuc }).ToList();

                    dataExport.SoBatThuong = lsBatThuong.Count();
                    string NgayBatThuong = "";
                    foreach (var item in lsBatThuong)
                    {
                        NgayBatThuong += item.DateVaoThuc.Date.ToString("MM/dd") + ", ";
                    }
                    dataExport.NgayBatThuong = NgayBatThuong;

                    // Xóa các dòng mà có số giờ  làm < 30 phút
                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        if (lsDataSheet[i].DateRa.Day != lsDataSheet[i].DateVao.Day || lsDataSheet[i].Time < 0.5)
                        {
                            lsDataSheet.RemoveAt(i);
                            i--;
                        }
                    }
                    // Kiểm tra các trường hợp quẹt thẻ để tính ra giờ làm và giờ tăng ca
                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        DateTime dateVao = lsDataSheet[i].DateVao;
                        DateTime dateRa = lsDataSheet[i].DateRa;

                        double gioLam = 0;
                        double tangca = 0;

                        if (dateRa.TimeOfDay < Moc1730.TimeOfDay)
                        {
                            var timeLam = dateRa.TimeOfDay - dateVao.TimeOfDay;
                            gioLam = timeLam.Hours + timeLam.Minutes / 60.0;
                        }
                        if (dateRa.TimeOfDay >= Moc1730.TimeOfDay)
                        {
                            if (dateVao.TimeOfDay < Moc1700.TimeOfDay)
                            {
                                var timeLam = Moc1700.TimeOfDay - dateVao.TimeOfDay;
                                gioLam = timeLam.Hours + timeLam.Minutes / 60.0;
                                var timeTangCa = dateRa.TimeOfDay - Moc1700.TimeOfDay;
                                tangca = timeTangCa.Hours + timeTangCa.Minutes / 60.0;
                            }
                            else
                            {
                                var timeTangCa = dateRa.TimeOfDay - dateVao.TimeOfDay;
                                tangca = timeTangCa.Hours + timeTangCa.Minutes / 60.0;
                            }

                        }

                        if (dateVao.TimeOfDay < Moc1130.TimeOfDay && dateRa.TimeOfDay > Moc1300.TimeOfDay)
                        {
                            gioLam -= 1.5;
                        }

                        lsDataSheet[i].GioLam = gioLam;
                        lsDataSheet[i].TangCa = tangca;
                    }

                    // Tính tổng số ngày làm và giờ tăng ca trong 1 tháng
                    double SumGioLam = lsDataSheet.Sum(r => r.GioLam);
                    double SumTangCa = lsDataSheet.Sum(r => r.TangCa);

                    double soNgayLam = (int)SumGioLam / 8;
                    double soGioLam = SumGioLam % 8;
                    double NgayTangCa = (int)SumTangCa / 8;
                    double gioTangCa = SumTangCa % 8;

                    dataExport.SoNgayLam = soNgayLam;
                    dataExport.SoGioLam = soGioLam;
                    dataExport.NgayTangCa = NgayTangCa;
                    dataExport.SoGioTangCa = gioTangCa;

                    // add dữ liệu của người đó vào list
                    lsExport.Add(dataExport);

                    DataTable aabb = ToDataTable(lsDataSheet);

                    var lsKhongDuNgay = (from data in lsDataSheet
                                         where data.GioLam != 8
                                         select new
                                         {
                                             DateVao = data.DateVaoThuc,
                                             DateRa = data.DateRaThuc,
                                             Time = data.GioLam
                                         }).ToList();

                    if (lsKhongDuNgay.Count != 0 || lsBatThuong.Count != 0)
                    {
                        ws.Cells[$"A{flag - 3}"].Value = Name;
                        ws.Cells[$"A{flag - 3}:C{flag - 3}"].Merge = true;
                        ws.Cells[$"A{flag - 3}"].Style.Font.Bold = true;
                        ws.Cells[$"A{flag - 3}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[$"A{flag - 3}"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        // vẽ Boder
                        ws.Cells[flag - 3, 1, flag - 3, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 3, 1, flag - 3, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 3, 1, flag - 3, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 3, 1, flag - 3, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }

                    if (lsKhongDuNgay.Count != 0)
                    {
                        ws.Cells[$"A{flag - 2}"].Value = "Ngày không đủ 8 tiếng";
                        ws.Cells[$"A{flag - 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[$"A{flag - 2}"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        ws.Cells[$"A{flag - 2}:C{flag - 2}"].Merge = true;

                        ws.Cells[$"A{flag - 1}"].Value = "Ngày vào";
                        ws.Cells[$"B{flag - 1}"].Value = "Ngày ra";
                        ws.Cells[$"C{flag - 1}"].Value = "Số giờ";

                        ws.Cells[$"A{flag}"].LoadFromCollection(lsKhongDuNgay, false);

                        // vẽ Boder
                        ws.Cells[flag - 2, 1, flag - 1 + lsKhongDuNgay.Count(), 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsKhongDuNgay.Count(), 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsKhongDuNgay.Count(), 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsKhongDuNgay.Count(), 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        flag += lsKhongDuNgay.Count() + 2;
                    }

                    if (lsBatThuong.Count != 0)
                    {
                        ws.Cells[$"A{flag - 2}"].Value = "Ngày bất thường";
                        ws.Cells[$"A{flag - 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[$"A{flag - 2}"].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);

                        ws.Cells[$"A{flag - 2}:C{flag - 2}"].Merge = true;

                        ws.Cells[$"A{flag - 1}"].Value = "Ngày vào";
                        ws.Cells[$"B{flag - 1}"].Value = "Ngày ra";
                        ws.Cells[$"C{flag - 1}"].Value = "Time";

                        ws.Cells[$"A{flag}"].LoadFromCollection(lsBatThuong, false);

                        // vẽ Boder
                        ws.Cells[flag - 2, 1, flag - 1 + lsBatThuong.Count(), 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsBatThuong.Count(), 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsBatThuong.Count(), 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[flag - 2, 1, flag - 1 + lsBatThuong.Count(), 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        flag += lsBatThuong.Count() + 2;
                    }

                    if (lsKhongDuNgay.Count != 0 || lsBatThuong.Count != 0)
                    {
                        flag += 3;
                    }



                }


                // string pathFile = Path.Combine(dialog.SelectedPath, $"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
                FileInfo excelFile = new FileInfo($"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
                pck.SaveAs(excelFile);
                Process.Start($"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
            }

            gridControl1.DataSource = lsExport;
        }
        class DataRaw
        {
            public DateTime DateVao { get; set; }
            public DateTime DateVaoThuc { get; set; }
            public DateTime DateRa { get; set; }
            public DateTime DateRaThuc { get; set; }

            public double Time { get; set; }
            public double GioLam { get; set; }
            public double TangCa { get; set; }
        }

        class DataExport
        {
            public string Name { get; set; }
            public string UserId { get; set; }
            public string CCCD { get; set; }
            public double SoNgayLam { get; set; }
            public double SoGioLam { get; set; }
            public double NgayTangCa { get; set; }
            public double SoGioTangCa { get; set; }
            public double SoBatThuong { get; set; }
            public string NgayBatThuong { get; set; }
        }
    }
}
