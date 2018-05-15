using Microsoft.VisualBasic;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Automatic_Folder_Generation_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string[] iniArray = new string[30];
        private int FolderCounter;
        private string result;
        private string Folder;
        private int DefaultCounter;

        private Regex r = new Regex("[\\x00-\\x1f<>:\"/\\\\|?*]"
                + "|^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9]|CLOCK\\$)(\\.|$)"
                + "|[\\. ]quot;", RegexOptions.IgnoreCase);

        // Win32APIの GetPrivateProfileString を使う宣言
        [DllImport("KERNEL32.DLL")]
        private static extern uint GetPrivateProfileString(
            string lpAppName,
            string lpKeyName,
            string lpDefault,
            StringBuilder lpReturnedString,
            uint nSize,
            string lpFileName
            );

        //セクション名を取得する関数
        [DllImport("kernel32", CharSet = CharSet.Auto)]
        private static extern int GetPrivateProfileString(
            string section,
            string key,
            string def,
            char[] retVal,
            int size,
            string filePath);

        //キー名を取得する関数
        [DllImport("KERNEL32.DLL", EntryPoint = "GetPrivateProfileStringA")]
        private static extern uint GetPrivateProfileStringByByteArray(string lpAppName,
            string lpKeyName,
            string lpDefault,
            byte[] lpReturnedString,
            uint nSize,
            string lpFileName);

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (DefaultSetting() == true)
                {
                    MaximizeBox = false;
                    MinimumSize = Size;
                    MaximumSize = Size;
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("エラーが発生しましたので、アプリを終了します。");
                Close();
            }
            CreateBtn.Enabled = true;
            //フォームが最小化されないようにする
            MaximizeBox = false;
            MinimumSize = Size;
            MaximumSize = Size;
        }

        private bool DefaultSetting()
        {
            iniArray[0] = "00_製作伝票";
            iniArray[1] = "01_見積もり";
            iniArray[2] = "02_客先資料";
            iniArray[3] = "03_提出資料";
            iniArray[4] = "04_仕様";
            iniArray[5] = "05_検討";
            iniArray[6] = "06_部材発注";
            iniArray[7] = "07_客先打合せ";
            iniArray[8] = "20_回路設計";
            iniArray[9] = "30_機構設計";
            iniArray[10] = "40_ソフト設計";
            iniArray[11] = "50_出荷";
            iniArray[12] = "60_稟議";
            iniArray[13] = "99_ワークフロー";
            for (int i = 0; i < 14; i++)
            {
                FolderList.Items.Add(iniArray[i]);
                FolderCounter++;
            }
            return true;
        }

        private void CreateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                CreateBtn.Text = "生成中・・・";
                CreateBtn.Enabled = false;
                Folder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString();//デスクトップのパスを取得
                if (check() == true)
                {
                    DirectoryInfo di = new DirectoryInfo(Folder + @"\" + result.Trim() + "(" + ClientName.Text.Trim() + " " + ProjectName.Text.Trim() + ")");
                    di.Create();
                    for (int i = 0; i < FolderCounter; i++)
                    {
                        di.CreateSubdirectory(iniArray[i]);
                    }

                    DirectoryInfo[] subFolders = di.GetDirectories("*", SearchOption.AllDirectories);
                    if (subFolders.Length == FolderCounter)
                    {
                        MessageBox.Show("正常にフォルダの生成が完了しました！");
                        di = null;//親フォルダのDirectoryInfoのインスタンスを破棄
                        subFolders = null;//子フォルダのインスタンスを破棄
                        CreateBtn.Enabled = true;
                        CreateBtn.Text = "自動生成開始";
                    }
                    else
                    {
                        MessageBox.Show("フォルダの自動生成に失敗しました。");
                        CreateBtn.Text = "自動生成開始";
                        di = null;//親フォルダのDirectoryInfoのインスタンスを破棄
                        subFolders = null;//子フォルダのインスタンスを破棄
                        CreateBtn.Enabled = true;
                        return;
                    }
                }
                else
                {
                    CreateBtn.Text = "自動生成開始";
                    CreateBtn.Enabled = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成処理エラー：" + ex.ToString());
                return;
            }
        }

        private bool check()
        {
            try
            {
                var Trimed = WorkNumber.Text.Trim();
                var TrimedProj = ProjectName.Text.Trim();
                var TrimedClien = ClientName.Text.Trim();
                //工番の長さが8文字の場合
                if (WorkNumber.Text.Length == 8)
                {
                    //工番に不正文字が入っていた場合
                    if (r.IsMatch(WorkNumber.Text))
                    {
                        MessageBox.Show("使用できない文字があります。"
                             + "\r\n" + "工番の入力内容を再度確認して下さい。");
                        return false;
                    }
                    //x or zが小文字だった場合、大文字に変換
                    if (Regex.IsMatch(WorkNumber.Text, @"^(\d{2}[x,z]-\d{4})$") == true)
                    {
                        result = Strings.StrConv(WorkNumber.Text, VbStrConv.Uppercase);
                    }
                    //X or Zが大文字で入力されていた時
                    else
                    {
                        result = WorkNumber.Text;
                    }
                    //工番名に不正文字が入っていた場合
                    if (r.IsMatch(Trimed))
                    {
                        MessageBox.Show("工番名に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }
                    //顧客名称に不正文字が入っていた場合
                    if (r.IsMatch(TrimedClien))
                    {
                        MessageBox.Show("顧客名称に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }
                    //案件名称に不正文字が入っていた場合
                    if (r.IsMatch(TrimedProj))
                    {
                        MessageBox.Show("案件名称に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }
                    //3項目のいずれかが空白の時
                    if (Trimed == string.Empty || TrimedProj == string.Empty || TrimedClien == string.Empty)
                    {
                        MessageBox.Show("工番・案件名称・顧客名称のいずれかが空欄になっていますので、入力して下さい。");
                        return false;
                    }
                    if (Regex.IsMatch(Trimed, @"\d{2}([X,Z]|[x,z])-\d{4}") == false)
                    {
                        MessageBox.Show("開発工番の形は、「◯◯（整数2つ）X（または　Z）-◯◯◯◯（整数４つ）」です。");
                        return false;
                    }
                    if ((result + "(" + TrimedClien + " " + TrimedProj + ")").Length > 100)
                    {
                        MessageBox.Show("フォルダの名称が長すぎます（最大100文字）" + "\r\n" + "フォルダ名を短くしてください");
                        return false;
                    }
                    DirectoryInfo di = new DirectoryInfo(Folder + @"\" + result + "(" + TrimedClien + " " + TrimedProj + ")");
                    //フォルダ重複確認
                    if (di.Exists == true)
                    {
                        MessageBox.Show("既に" + result + "(" + TrimedClien + " " + TrimedProj + ")" + "は存在します。");
                        return false;
                    }
                }
                else
                {
                    //12X(x)1234 or 12Z(z)1234の場合（ハイフン抜け）
                    if (Regex.IsMatch(Trimed, @"^(\d{2}([X,Z]|[x,z])\d{4})$") == true)

                    {
                        MessageBox.Show("ハイフンは、必ず追加して下さい");
                        return false;
                    }
                    //X(x) or Z(z)がついていない時
                    if (Regex.IsMatch(Trimed, @"^(\d{2}-\d{4})$") == true)
                    {
                        MessageBox.Show("社外品の場合は、開発年度の後ろに'X'を、自社製品の場合は'Z'を必ず付加して下さい");
                        return false;
                    }
                    //例：X(x) or Z(z)単独
                    if (Regex.IsMatch(Trimed, @"^(([X,Z]|[x,z])|\d{0,1}([X,Z]|[x,z])-\d{0,4})$") == true)
                    {
                        MessageBox.Show("工番を、例に習って修正して下さい");
                        return false;
                    }
                    //12X(x) or 12Z(z)の時
                    if (Regex.IsMatch(Trimed, @"^(\d{0,2}([X,Z]|[x,z]))$") == true)
                    {
                        MessageBox.Show("通し番号は、必ず追加して下さい");
                        return false;
                    }
                    //12X(x)- or 12Z(z)-の時
                    if (Regex.IsMatch(Trimed, @"^(\d{0,2}([X,Z]|[x,z])-)$") == true)
                    {
                        MessageBox.Show("通し番号は必ず4桁で表現して下さい。通し番号が0番だった場合は、0000と表現して下さい。");
                        return false;
                    }
                    //例：通し番号の桁数が足りない時
                    if (Regex.IsMatch(Trimed, @"^(\d{2}([X,Z]|[x,z])-\d{0,3})$") == true)
                    {
                        MessageBox.Show("通し番号は必ず4桁で表現して下さい。通し番号が0番だった場合は、0000と表現して下さい。");
                        return false;
                    }
                    //3項目のいずれかが空白の時
                    if (Trimed == string.Empty || TrimedProj == string.Empty || TrimedClien == string.Empty)
                    {
                        MessageBox.Show("工番・案件名称・顧客名称のいずれかが空欄になっていますので、入力して下さい。");
                        return false;
                    }
                    if (Regex.IsMatch(Trimed, @"\d{2}([X,Z]|[x,z])-\d{4}") == false)
                    {
                        MessageBox.Show("開発工番の形は、「◯◯（整数2つ）X（または　Z）-◯◯◯◯（整数４つ）」です。");
                        return false;
                    }

                    //工番名に不正文字が入っていた場合
                    if (r.IsMatch(Trimed))
                    {
                        MessageBox.Show("工番名に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }

                    //案件名称に不正文字が入っていた場合
                    if (r.IsMatch(TrimedProj))
                    {
                        MessageBox.Show("案件名称に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }
                    //顧客名称に不正文字が入っていた場合
                    if (r.IsMatch(TrimedClien))
                    {
                        MessageBox.Show("顧客名称に、使用できない文字があります。"
                             + "\r\n" + "案件名称の入力内容を再度確認して下さい。");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("入力チェックエラー：" + ex.ToString());
                return false;
            }
            return true;
        }

        private void iniファイルを読み込むToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                ofd.Filter = "構成ファイル（*.ini）|*.ini";
                ofd.Title = "読み込みたいiniファイルを選択して下さい。";
                ofd.RestoreDirectory = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    var iniFileName = ofd.FileName;
                    FolderList.Items.Clear();
                    FolderCounter = 0;
                    int iniLines = (File.ReadAllLines(iniFileName).Length) - 1;
                    //セクション名取得
                    char[] buf = new char[1024];
                    int ret = GetPrivateProfileString(null, null, null, buf, buf.Length, iniFileName);
                    string sectionsText = new string(buf, 0, ret);
                    string[] sections = sectionsText.TrimEnd('\0').Split('\0');
                    //セクション名が存在しない時
                    for (int i = 0; i < sections.Length; i++)
                    {
                        sections[i] = Strings.StrConv(sections[i], VbStrConv.Narrow);
                        if (sections[i] == "")
                        {
                            MessageBox.Show("セクションが存在しないものがあります。" + "\r\n" + "iniファイルを確認して下さい。");
                            Close();
                            return;
                        }
                        else if (sections[0] != "FolderName")
                        {
                            MessageBox.Show("セクションが存在しないものがあります。" + "\r\n" + "iniファイルを確認して下さい。");
                            Close();
                            return;
                        }
                        else
                        {
                            ;
                        }
                    }

                    for (int i = 0; i < sections.Length; i++)
                    {
                        for (int j = 0; j < sections.Length; j++)
                        {
                            if (i != j)
                            {
                                if (sections[i] == sections[j])
                                {
                                    MessageBox.Show("セクション名で、重複が見つかりました" + "\r\n"
                                        + "セクション名：" + sections[j] + "\r\n" + "iniファイルを確認して下さい。");
                                    Close();
                                    return;
                                }
                            }
                        }
                    }
                    //キー名取得
                    byte[] ar1 = new byte[1024];
                    uint resultSize1 = GetPrivateProfileStringByByteArray("FolderName", null, "default", ar1, (uint)ar1.Length, iniFileName);
                    string result1 = Encoding.Default.GetString(ar1, 0, (int)resultSize1 - 1);
                    string[] keys = result1.Split('\0');
                    int KeyCnt = 0;
                    for (int i = 0; i < keys.Length; i++)
                    {
                        KeyCnt = 0;
                        for (int j = 0; j < keys.Length; j++)
                        {
                            if (Regex.IsMatch(keys[j], @"^(([0-9]|1[0-9]|2[0-8](.[0-9]+)?)|29)$") == false)
                            {
                                MessageBox.Show("キー名に、不正なキーが使用されました" + "\r\n" + "該当のキー名："
                                    + keys[j] + "\r\n" + "キー名は0 ~ 29までを使用して下さい" + "\r\n"
                                    + "全角で入力している場合は半角で入力し直して下さい。");
                                Close();
                                return;
                            }
                            if (i != j)
                            {
                                if (keys[i] == keys[j])
                                {
                                    MessageBox.Show("キー名で重複が見つかりました" + "\r\n" + "キー名：" + keys[j] + "\r\n"
                                        + "iniファイルを確認して下さい。");
                                    Close();
                                    return;
                                }
                            }
                            if (keys[j] == "")
                            {
                                MessageBox.Show("キー名が不正なiniファイルです。" + "iniファイルのキー名とバリュー名の書式は" + "\r\n"
                                    + "Key = Valueの形で、記述して下さい");
                                Close();
                                return;
                            }
                            if (keys[j] != KeyCnt.ToString())
                            {
                                MessageBox.Show("キー名が不正です。" + "キー名：" + keys[j]
                                    + "\r\n" + "必ずキー名は連番で振り分けて下さい");
                                Close();
                                return;
                            }
                            KeyCnt++;
                        }
                    }

                    // 指定ファイルのバリューの一覧を得る
                    StringBuilder builder = new StringBuilder(1024);

                    for (int i = 0; i < iniArray.Length; i++)
                    {
                        GetPrivateProfileString(
                            sections[0],      // セクション名
                            i.ToString(),          // キー名
                            "null",   // 値が取得できなかった場合に返される初期値
                            builder, // 格納先
                            Convert.ToUInt32(builder.Capacity), iniFileName);   // iniファイル名
                        if (r.IsMatch(builder.ToString()))
                        {
                            MessageBox.Show("バリューの中に、使用できない文字を含んでいます。"
                                 + "\r\n" + "iniファイルの情報を確認の上、再度実行して下さい。");
                            Close();
                            return;
                        }
                        if (builder.ToString() == "null")
                        {
                            break;
                        }
                        if (Regex.IsMatch(builder.ToString(), @"^(.*=.*)$"))
                        {
                            MessageBox.Show("入力ミスの可能性があります" + "バリュー名：" + builder.ToString() + "\r\n" + "iniファイルを確認して下さい。");
                            Close();
                            return;
                        }
                        iniArray[i] = builder.ToString();
                        if (iniArray[i] == "")
                        {
                            MessageBox.Show("値が設定されていないものがあります。" + "キー名：" + i.ToString()
                                + "\r\n" + "iniファイルを確認して下さい。");
                            Close();
                            return;
                        }
                        if (iniArray[i].Length > 60)
                        {
                            MessageBox.Show("バリュー名が60文字を超えているため、読み込みに失敗しました。");
                            Close();
                            return;
                        }

                        FolderCounter++;
                        if (i != 0)
                        {
                            for (int j = 1; j < FolderCounter; j++)
                            {
                                if (iniArray[j - 1] == builder.ToString())
                                {
                                    MessageBox.Show("iniファイルで重複が見つかりました。" + "\r\n" + "Value：" + builder.ToString());
                                    Close();
                                    return;
                                }
                            }
                        }
                    }
                    //ここで、リストボックスにフォルダを列挙
                    for (int k = 0; k < FolderCounter; k++)
                    {
                        {
                            FolderList.Items.Add(iniArray[k]);
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("エラー発生：" + "iniファイルを確認して下さい");
                Close();
                return;
            }
        }
    }
}