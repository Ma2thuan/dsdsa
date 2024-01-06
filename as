using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LAB04
{
    public partial class Form1 : Form
    {
        LAB04Entities db = new LAB04Entities();
        public Form1()
        {
            InitializeComponent();
            btn_Luu.Enabled = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có muốn đóng chương trình hay không?", "Xác nhận", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void rad_ToaSan_CheckedChanged(object sender, EventArgs e)
        {
            lbl_PhuCapTangGio.Text = "Số giờ làm thêm:";
        }

        private void rad_ThuongTru_CheckedChanged(object sender, EventArgs e)
        {
            lbl_PhuCapTangGio.Text = "Phụ cấp:";
        }

        private void btn_Them_Click(object sender, EventArgs e)
        {
            txt_MaPV.Clear();
            txt_Ten.Clear();
            txt_DienThoai.Clear();
            dtp_NgayVaoLam.ResetText();
            txt_PhuCapTangGio.Clear();
            rad_Nam.Checked = true;
            rad_ToaSoan.Checked = true;
            txt_MaPV.Focus();
        }

        private void btn_Luu_Click(object sender, EventArgs e)
        {
            db.PHONGVIENs.Add(new PHONGVIEN()
            {
                MAPV = txt_MaPV.Text,
                HOTEN = txt_Ten.Text,
                GIOITINH = (rad_Nam.Checked ? "Nam" : "Nữ"),
                SDT = txt_DienThoai.Text,
                NGAYVAOLAM = dtp_NgayVaoLam.Value.Date,              
                LOAIPV = (rad_ToaSoan.Checked ? "Tòa Soạn" : "Thường Trú"),
                OT = (rad_ToaSoan.Checked ? int.Parse(txt_PhuCapTangGio.Text) : 0),
                PHUCAP = (rad_ThuongTru.Checked ? float.Parse(txt_PhuCapTangGio.Text) : 0)

            });

            db.SaveChanges();
            Form1_Load(sender, e);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var dbLoad = new LAB04Entities();
            var PhongVien = dbLoad.PHONGVIENs.ToList();

            lv_ThongTinPhongVien.Items.Clear();
            lv_ThongTinPhongVien.Columns.Clear();

            lv_ThongTinPhongVien.Columns.Add("Mã PV");
            lv_ThongTinPhongVien.Columns.Add("Tên");
            lv_ThongTinPhongVien.Columns.Add("Giới Tính");
            lv_ThongTinPhongVien.Columns.Add("Ngày Vào Làm");
            foreach (var pv in PhongVien)
            {
                ListViewItem item = new ListViewItem(pv.MAPV);
                item.SubItems.Add(pv.HOTEN);
                
                item.SubItems.Add(pv.GIOITINH);
                item.SubItems.Add(pv.NGAYVAOLAM.ToString());

                if (pv.NGAYVAOLAM <= DateTime.Now.AddDays(-1827))
                {
                    // Nếu ngày vào làm lớn hơn 5 năm tô màu dé lồ
                    item.BackColor = Color.Yellow;
                }

                item.Tag = pv;


                lv_ThongTinPhongVien.Items.Add(item);
            }
        }

        private void txt_MaPV_TextChanged(object sender, EventArgs e)
        {
            if (txt_MaPV.Text != "" && txt_Ten.Text != "" && txt_PhuCapTangGio.Text != "" && txt_DienThoai.Text != "")
            {
                btn_Luu.Enabled = true;
            }
            else
                btn_Luu.Enabled = false;
        }

        private void txt_Ten_TextChanged(object sender, EventArgs e)
        {
            if (txt_MaPV.Text != "" && txt_Ten.Text != "" && txt_PhuCapTangGio.Text != "" && txt_DienThoai.Text != "")
            {
                btn_Luu.Enabled = true;
            }
            else
                btn_Luu.Enabled = false;
        }

        private void txt_DienThoai_TextChanged(object sender, EventArgs e)
        {
            if (txt_MaPV.Text != "" && txt_Ten.Text != "" && txt_PhuCapTangGio.Text != "" && txt_DienThoai.Text != "")
            {
                btn_Luu.Enabled = true;
            }
            else
                btn_Luu.Enabled = false;
        }

        private void txt_PhuCapTangGio_TextChanged(object sender, EventArgs e)
        {
            if (txt_MaPV.Text != "" && txt_Ten.Text != "" && txt_PhuCapTangGio.Text != "" && txt_DienThoai.Text != "")
            {
                btn_Luu.Enabled = true;
            }
            else
                btn_Luu.Enabled = false;
        }

        private void lv_ThongTinPhongVien_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lv_ThongTinPhongVien.SelectedItems.Count > 0)
            {
                ListViewItem selectItem = lv_ThongTinPhongVien.SelectedItems[0];

                var PhongVien = (PHONGVIEN)selectItem.Tag;

                txt_MaPV.Text = PhongVien.MAPV;
                txt_Ten.Text = PhongVien.HOTEN;

                if (PhongVien.GIOITINH == "Nam")
                    rad_Nam.Checked = true;
                else
                    rad_Nu.Checked = true;

                dtp_NgayVaoLam.Value = DateTime.Parse(PhongVien.NGAYVAOLAM.ToString());

                txt_DienThoai.Text = PhongVien.SDT;


                if (PhongVien.LOAIPV == "Tòa Soạn")
                {
                    rad_ToaSoan.Checked = true;
                    txt_PhuCapTangGio.Text = (PhongVien.OT).ToString();
                }
                else
                {
                    rad_ThuongTru.Checked = true;
                    txt_PhuCapTangGio.Text = (PhongVien.PHUCAP).ToString();
                }



            }
        }

        private void btn_Xoa_Click(object sender, EventArgs e)
        {
            if (lv_ThongTinPhongVien.SelectedItems.Count > 0)
            {
                //Lấy 
                ListViewItem chosenItem = lv_ThongTinPhongVien.SelectedItems[0];

                var PhongVien = (PHONGVIEN)chosenItem.Tag;

                var entity = db.PHONGVIENs.Find(PhongVien.MAPV);

                DialogResult dialogResult = MessageBox.Show("Bạn có muốn xóa phóng viên?", "Xác nhận xóa", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes && entity != null)
                {

                    // Loại bỏ đơn ra database
                    db.PHONGVIENs.Remove(entity);

                    //Lưu thay đổi
                    db.SaveChanges();

                    //Loại bỏ ra khỏi list view
                    chosenItem.Remove();
                }
            }
        }

        private void btn_Sua_Click(object sender, EventArgs e)
        {
            if (lv_ThongTinPhongVien.SelectedItems.Count > 0)
            {
                ListViewItem chosenItem = lv_ThongTinPhongVien.SelectedItems[0];

                var PhongVien = (PHONGVIEN)chosenItem.Tag;

                db.PHONGVIENs.Attach(PhongVien);

                //Update database
                if (PhongVien.LOAIPV == "Tòa Soạn")
                {
                    PhongVien.OT = int.Parse(txt_PhuCapTangGio.Text); 
                }
                else
                    PhongVien.PHUCAP = int.Parse(txt_PhuCapTangGio.Text);

                // save
                db.SaveChanges();

                //update lv

                if (PhongVien.LOAIPV == "Tòa Soạn")
                {
                    txt_PhuCapTangGio.Text = PhongVien.OT.ToString();
                }
                else
                    txt_PhuCapTangGio.Text = PhongVien.PHUCAP.ToString();


                db.SaveChanges();

                Form1_Load(sender, e);
            }
        }

        private void btn_SapXep_Click(object sender, EventArgs e)
        {
            lv_ThongTinPhongVien.ListViewItemSorter = new ListViewDateComparer();
            lv_ThongTinPhongVien.Sort();
        }

        public class ListViewDateComparer : IComparer
        {
            public int Compare(object x, object y)
            {
                ListViewItem itemX = (ListViewItem)x;
                ListViewItem itemY = (ListViewItem)y;

                DateTime dateTimeX = DateTime.Parse(itemX.SubItems[2].Text);
                DateTime dateTimeY = DateTime.Parse(itemY.SubItems[2].Text);

                return -DateTime.Compare(dateTimeX, dateTimeY);
            }
        }

        private void btn_ThongKe_Click(object sender, EventArgs e)
        {
            int ts = db.PHONGVIENs.Count(pv => pv.LOAIPV == "Tòa Soạn");
            int tt = db.PHONGVIENs.Count(pv => pv.LOAIPV == "Thường Trú");


            MessageBox.Show("Số lượng PV tòa soạn : " + ts + '\n' +
                            "Số lượng PV thường trú: " + tt);
        }

        private void btn_Excel_Click(object sender, EventArgs e)
        {
            // Tạo ra 1 instance excel
            Excel.Application excelApp = new Excel.Application();

            // Tạo ra file chính (Workbook) trong Excel vừa mới tạo 
            Excel.Workbook excelWb = excelApp.Workbooks.Add
            (Excel.XlWBATemplate.xlWBATWorksheet); // Trong Workbook vừa tạo có chứa 1 sheet

            //  Sử dụng worksheet đầu tiên vừa được tạo ở phía trên
            Excel.Worksheet excelWs = excelWb.Worksheets[1];

            // Bắt đầu tại ô [1,1]
            Excel.Range excelRange = excelWs.Cells[1,1];
            // Thiết lập phong chữ
            excelRange.Font.Size = 16;
            excelRange.Font.Bold = true;
            excelRange.Font.Color = Color.Blue;
            excelRange.Value = "DANH SÁCH CÁC PHÓNG VIÊN TÁC NGHIỆP";


            // Lấy DS phóng viên
            var PhongVien = db.PHONGVIENs.Select(pv => new {Code = pv.MAPV , Name = pv.HOTEN , Date = pv.NGAYVAOLAM , pv.SDT }).ToList();

            int row = 3;

            foreach (var p in PhongVien)
            {
                excelWs.Range["A" + row].Value = p.Code;
                excelWs.Range["B" + row].Value = p.Name;
                excelWs.Range["C" + row].Value = p.Date;
                excelWs.Range["D" + row].ColumnWidth = 90;
                excelWs.Range["D" + row].Value = p.SDT;
                row++;
            }

            excelWs.Name = "Danh Sách Phóng Viên";    
            // Kích hoạt sheet nếu khi ta ghi dữ liệu mà ko nói rõ sheet nào được sử dụng thì mặc định dữ liệu sẽ được ghi trên trang của Excel
            excelWs.Activate();

            // Lưu File
            SaveFileDialog saveFileDialog = new SaveFileDialog();   
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                excelWs.SaveAs(saveFileDialog.FileName);
            excelApp.Quit();
        }

        private void btn_PDF_Click(object sender, EventArgs e)
        {
            // Chuẩn bị nguồn dữ liệu
            var data = db.PHONGVIENs.Select(p => new { MAPV = p.MAPV, HOTEN = p.HOTEN, NGAYVAOLAM = p.NGAYVAOLAM, SDT = p.SDT }).ToList();

            // Gán nguồn dữ liệu cho Crystal Report
            CrystalReport1 rpt = new CrystalReport1();
            rpt.SetDataSource(data);

            //Hiện thị báo cáo  
            Form2 frpt = new Form2();
            frpt.crystalReportViewer1.ReportSource = rpt;
            frpt.ShowDialog();
        }
    }

}
