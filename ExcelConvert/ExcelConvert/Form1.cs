using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelConvert
{
    public partial class ExcelConvert : Form
    {
        string gFilePath = "";
        public ExcelConvert()
        {
            InitializeComponent();

            dataGridView1.AllowDrop = true;

            //this.dataGridView1.DragDrop += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragDrop);
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            gFilePath = "";

            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effect = DragDropEffects.Copy;
                string[] filepath = (string[]) e.Data.GetData(DataFormats.FileDrop);
                gFilePath = filepath[0];

                //MessageBox.Show(System.Windows.Forms.Application.StartupPath);
                if (!gFilePath.Equals("") && true)
                {
                    fExcelLoad();
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

            //MessageBox.Show(gFilePath);

        }

        private void fExcelSave()
        {
            Microsoft.Office.Interop.Excel.Application application = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;

            try
            {
                string sRootPath = System.Windows.Forms.Application.StartupPath;

                application = new Microsoft.Office.Interop.Excel.Application();
                workbook = application.Workbooks.Open(sRootPath + "\\" + "주문.xls");
                worksheet = workbook.Worksheets.get_Item(1);

                application.Visible = false;

                range = worksheet.UsedRange;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for(int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        range[i + 2, j + 1] = String.IsNullOrEmpty((string) dataGridView1.Rows[i].Cells[j].Value) ? "" : dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }

                }

                worksheet.SaveAs(@"C:\Users\Kwon\Desktop\" + "주문_" + DateTime.Now.ToString("yyyyMMddhhmmss"));

                DeleteObject(worksheet);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
            catch(Exception ex)
            {

            }
        }

        private void fExcelLoad()
        {
            Microsoft.Office.Interop.Excel.Application application = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;

            try
            {
                application = new Microsoft.Office.Interop.Excel.Application();
                workbook = application.Workbooks.Open(gFilePath);
                worksheet = workbook.Worksheets.get_Item(1);

                application.Visible = false;

                range = worksheet.UsedRange;

                dataGridView1.Rows.Clear();

                int idx = -1;

                for (int i = 2; i <= range.Rows.Count; ++i)
                {
                    dataGridView1.Rows.Add();
                    idx++;

                    dataGridView1.Rows[idx].Cells[1].Value = "항공(인천)"; //상세주소
                    dataGridView1.Rows[idx].Cells[7].Value = "개인"; //용도구분
                    dataGridView1.Rows[idx].Cells[10].Value = "1"; //상세주소
                    dataGridView1.Rows[idx].Cells[23].Value = "115"; //상세주소

                    for (int j = 1; j <= range.Columns.Count; ++j)
                    {
                        switch (j)
                        {
                            case 25: //구매자
                                dataGridView1.Rows[idx].Cells[2].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            case 36: //개인통관번호
                                dataGridView1.Rows[idx].Cells[4].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            case 37: //구매자 핸드폰번호
                                dataGridView1.Rows[idx].Cells[6].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            case 29: //우편번호
                                dataGridView1.Rows[idx].Cells[8].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            case 30: //주소
                                dataGridView1.Rows[idx].Cells[9].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            case 23: //구매수량
                                dataGridView1.Rows[idx].Cells[17].Value = (range.Cells[i, j] as Range).Value2.ToString();
                                break;
                            default:
                                break;
                        }

                    }
                }

                //range.Cells[3, 3] = "123";
                //worksheet.SaveAs(@"C:\Users\Kwon\Desktop\123.xlsx");

                DeleteObject(worksheet);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
            catch(Exception ex)
            {

            }
            finally
            {

            }
            
        }

        private void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void btnExcelSave_Click(object sender, EventArgs e)
        {
            fExcelSave();
        }
    }
}
