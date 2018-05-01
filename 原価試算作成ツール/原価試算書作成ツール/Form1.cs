using System;
using System.IO;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.Linq;

namespace 原価試算書作成ツール
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string[] ExcelInfo = new string[10];
        private string[] RepNum = new string[10];
        private string DeskPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

        private Regex r =
        new Regex("[\\x00-\\x1f<>:\"/\\\\|?*]" + "|^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9]|CLOCK\\$)(\\.|$)" + "|[\\. ]quot;", RegexOptions.IgnoreCase);

        private void Form1_Load(object sender, EventArgs e)
        {
            Output.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            RepNum[0] = "-A001";
            RepNum[1] = "-C001";
            RepNum[2] = "-Q1";
            RepNum[3] = "-C101";
        }

        private void Dialog1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = @"\\Gserver\002_各種物件";
            ofd.RestoreDirectory = true;
            ofd.Title = "見積もり情報シートを選択して下さい";
            ofd.Filter = "Excel 97-2003形式(*.xls)|*.xls|Excelワークシート(*.xlsx)|*.xlsx";
            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(ofd.FileName);
                Input.Text = ofd.FileName;
            }
            ofd.Dispose();
        }

        private void Dialog2Btn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "フォルダを指定してください。";
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                //選択されたフォルダを表示する
                Console.WriteLine(fbd.SelectedPath);
                Output.Text = fbd.SelectedPath;
            }
            fbd.Dispose();
        }

        private void CreateBtn_Click(object sender, EventArgs e)
        {
            if (check() == true)
                if (Copy() == true)
                {
                    Console.WriteLine("コピーに成功");
                    if (ReadEx() == true)
                    {
                        var di = new DirectoryInfo(DeskPath + @"\原価試算書フォルダ");
                        var Num = 1;
                        if (File.Exists(DeskPath + @"\原価試算書フォルダ\【" + ExcelInfo[0] + " " + "-C00" + Num + "】" + "製造原価試算書.xlsx") == true)
                        {
                            File.Move(DeskPath + @"\原価試算書フォルダ\HR40-C001_製造原価試算書_T.xlsx",
                            DeskPath + @"\原価試算書フォルダ\【" + ExcelInfo[0] + " " + "-C00" + Num++ + "】" + "製造原価試算書.xlsx");
                        }
                        if (WriteExManu() == true)
                        {
                            File.Move(DeskPath + @"\原価試算書フォルダ\HR40-C001_製造原価試算書_T.xlsx",
                            DeskPath + @"\原価試算書フォルダ\【" + ExcelInfo[0] + " " + "-C001" + "】" + "製造原価試算書.xlsx");
                        }
                        else
                        {
                            MessageBox.Show("製造原価試算書の作成に失敗しました。");
                            return;
                        }
                        if (WriteExDev() == true)
                        {
                            File.Move(DeskPath + @"\原価試算書フォルダ\HR209-C101_開発原価試算書_T.xlsx",
                            DeskPath + @"\原価試算書フォルダ\【" + ExcelInfo[0] + " " + "-C101" + "】" + "開発原価試算書.xlsx");
                            MessageBox.Show("原価試算書の作成が完了しました。");
                        }
                        else
                        {
                            MessageBox.Show("開発原価試算書の作成に失敗しました。");
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine("見積もり情報シートからデータを読み込むのに失敗しました");
                        return;
                    }
                }
                else
                {
                    return;
                }
            File.Delete(DeskPath + @"\原価試算書フォルダ\HR40-C001_製造原価試算書_T.xlsx");
            File.Delete(DeskPath + @"\原価試算書フォルダ\HR209-C101_開発原価試算書_T.xlsx");
        }

        private bool Copy()
        {
            try
            {
                string[] dirs = Directory.GetFiles(@"\\Gserver\999_個人フォルダ\柴田\原紙テスト\", "【原紙】HR40-C001*.xlsx1");
                File.Copy(dirs[0], DeskPath + @"\HR40-C001_製造原価試算書.xlsx");

                string[] dirs1 = Directory.GetFiles(@"\\Gserver\999_個人フォルダ\柴田\原紙テスト\", "【原紙】HR209-C101*.xlsx");
                File.Copy(dirs1[0], DeskPath + @"\HR209-C101_開発原価試算書.xlsx");

                DirectoryInfo di = new DirectoryInfo(DeskPath + @"\原価試算書フォルダ");
                di.Create();
                File.Move(DeskPath + @"\HR40-C001_製造原価試算書.xlsx",
                    DeskPath + @"\原価試算書フォルダ\HR40-C001_製造原価試算書.xlsx");
                File.Move(DeskPath + @"\HR209-C101_開発原価試算書.xlsx",
                    DeskPath + @"\原価試算書フォルダ\HR209-C101_開発原価試算書.xlsx");
                di = null;
            }
            catch (Exception ex)
            {
                if (ex is IOException)
                {
                    MessageBox.Show("既に原価試算書フォルダは存在します");

                    return false;
                }
                if (ex is UnauthorizedAccessException)
                {
                    MessageBox.Show("ファイルへのアクセスが拒否されました");
                    return false;
                }
            }

            return true;
        }

        private bool check()
        {
            if (Input.Text.Trim() == "")
            {
                MessageBox.Show("見積もり情報シートが未選択です");
                return false;
            }
            if (Output.Text.Trim() == "")
            {
                MessageBox.Show("保存先が未選択です");
                return false;
            }
            if (Directory.Exists(Input.Text) == false)
            {
                MessageBox.Show("選択した場所は存在しません");
                return false;
            }
            if (Directory.Exists(Output.Text) == false)
            {
                MessageBox.Show("選択した保存場所は存在しません");
                return false;
            }

            return true;
        }

        private bool ReadEx()//セルの情報読み取り
        {
            IRow row; ICell cell;
            //var FilePath = Input.Text;
            var FilePath = @"\\Gserver\999_個人フォルダ\柴田\原価試算作成ツール\課題\RFP18041301_見積り情報シート(西部電機様_バルコン本体).xls";
            var work = WorkbookFactory.Create(FilePath);
            var Sheet = work.GetSheetAt(0);

            row = Sheet.GetRow(0);//行
            cell = row.GetCell(11);
            string[] sp = cell.ToString().Split('：');//開発受付番号を全角「：」から分割
            ExcelInfo[0] = sp[1].Substring(0, 8);//開発工番のみ

            row = Sheet.GetRow(8);
            cell = row.GetCell(1);
            ExcelInfo[1] = cell.ToString();//件名

            row = Sheet.GetRow(9);
            cell = row.GetCell(1);
            ExcelInfo[2] = cell.ToString();//顧客名

            row = Sheet.GetRow(7);
            cell = row.GetCell(1);
            ExcelInfo[3] = "●";//見積もり種別（■）

            row = Sheet.GetRow(7);
            cell = row.GetCell(3);
            ExcelInfo[4] = "○";//見積もり種別（□）

            ExcelInfo[5] = "Ver.1";//バージョン

            row = Sheet.GetRow(37);//行
            cell = row.GetCell(3);
            ExcelInfo[6] = cell.ToString();//営業想定単価（想定単価）

            row = Sheet.GetRow(36);//行
            cell = row.GetCell(1);
            sp = cell.ToString().Split('想');//「想」から分割
            ExcelInfo[7] = sp[0];//開発費

            return true;
        }

        private bool WriteExManu()//製造原価試算書の書き込み用関数
        {
            try
            {
                var WritePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\原価試算書フォルダ\HR40-C001_製造原価試算書.xlsx";
                var WWork = WorkbookFactory.Create(WritePath);
                var WSheet = WWork.GetSheetAt(0);
                var row = WSheet.GetRow(12);
                var cell = row.GetCell(7);
                cell.SetCellValue(ExcelInfo[0] + RepNum[2]);// -Q1

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(12);
                cell = row.GetCell(4);
                cell.SetCellValue(ExcelInfo[0] + RepNum[0]);// -A001

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(13);
                cell = row.GetCell(6);
                cell.SetCellValue(ExcelInfo[0] + RepNum[1]);// -C001

                WSheet = WWork.GetSheetAt(0);//製品名
                row = WSheet.GetRow(13);
                cell = row.GetCell(2);
                cell.CellStyle.ShrinkToFit = true;//全体を縮小して表示
                cell.SetCellValue(ExcelInfo[1]);

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(15);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[2]);//顧客名

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(15);
                cell = row.GetCell(6);
                cell.SetCellValue(ExcelInfo[5]);//バージョン

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(20);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[3]);//試算区分「■」

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(21);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[4]);//試算区分「□」

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(20);
                cell = row.GetCell(9);
                cell.SetCellValue(ExcelInfo[6]);//営業想定単価

                using (var fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\原価試算書フォルダ\HR40-C001_製造原価試算書_T.xlsx", FileMode.Create))
                {
                    WWork.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
            return true;
        }

        private bool WriteExDev()//開発原価試算書の書き込み用関数
        {
            try
            {
                var WritePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\原価試算書フォルダ\HR209-C101_開発原価試算書.xlsx";
                var WWork = WorkbookFactory.Create(WritePath);
                var WSheet = WWork.GetSheetAt(0);
                var row = WSheet.GetRow(12);
                var cell = row.GetCell(7);
                cell.SetCellValue(ExcelInfo[0] + RepNum[2]);// -Q1

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(12);
                cell = row.GetCell(4);
                cell.SetCellValue(ExcelInfo[0] + RepNum[0]);// -A001

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(13);
                cell = row.GetCell(6);
                cell.SetCellValue(ExcelInfo[0] + RepNum[3]);// -C101

                WSheet = WWork.GetSheetAt(0);//製品名
                row = WSheet.GetRow(13);
                cell = row.GetCell(2);
                cell.CellStyle.ShrinkToFit = true;
                cell.SetCellValue(ExcelInfo[1]);

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(15);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[2]);//顧客名

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(15);
                cell = row.GetCell(6);
                cell.SetCellValue(ExcelInfo[5]);//バージョン

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(20);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[3]);//試算区分「■」

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(21);
                cell = row.GetCell(2);
                cell.SetCellValue(ExcelInfo[4]);//試算区分「□」

                WSheet = WWork.GetSheetAt(0);
                row = WSheet.GetRow(20);
                cell = row.GetCell(9);
                cell.SetCellValue(ExcelInfo[7]);//営業想定単価
                using (var fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\原価試算書フォルダ\HR209-C101_開発原価試算書_T.xlsx", FileMode.Create))
                {
                    WWork.Write(fs);
                }
            }
            catch (Exception ex)
            {
                if (ex is IOException)
                {
                    MessageBox.Show("書き込み中にエラーが発生しました");
                    return false;
                }
                else if (ex is FileNotFoundException)
                {
                    MessageBox.Show("書き込み先のExcelシートが見つかりませんでした。");
                    return false;
                }
                else
                {
                    MessageBox.Show("不明なエラーが発生しました");
                    return false;
                }
            }
            return true;
        }
    }
}