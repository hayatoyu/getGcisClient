using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace getGcisClient
{
    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern void FreeConsole();

        enum QueryType
        {
            Server,Local
        }
        
        private class ServerList
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public override string ToString()
            {
                return Value;
            }
        }

        private ServerList serverList = null;

        public Form1()
        {
            InitializeComponent();            
            cb_ServerList.Items.Add(new ServerList { Name = "測試用(10.10.8.30)", Value = "10.10.8.30" });
            cb_ServerList.Items.Add(new ServerList { Name = "正式用(10.1.199.146)",Value = "10.1.199.146" });
            cb_ServerList.DisplayMember = "Name";
            
            txt_Port.Text = ConfigurationManager.AppSettings["port"];
            AllocConsole();     // 這行的意思是我把程式中 Console 輸出的訊息，都顯示到 Console 的視窗中。
            //Console.Beep();
            Console.WriteLine("===== 歡迎使用商工行政開放資料服務 =====");
            Console.WriteLine("注意：請不要關閉此視窗");       // 用了 AllocConsole() 後，關掉 Console 視窗程式也會終止。
        }

        private void btn_SelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select .xls File";
            dialog.Filter = "xls files (*.xls;*.xlsx)|*.xls;*.xlsx";
            dialog.ShowDialog();
            txt_FilePath.Text = dialog.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            txt_SaveFolder.Text = dialog.SelectedPath;
        }

        private void btn_Connect_Click(object sender, EventArgs e)
        {
            // 先檢查各項設定有沒有設
            if(ValidateConnect(QueryType.Server))
            {
                Console.WriteLine("資料驗證成功...準備連線至 {0}，連接埠 {1}...",serverList.Value,txt_Port.Text);
                

                Client client = new Client(serverList.Value, txt_Port.Text, txt_FilePath.Text, txt_SaveFolder.Text, this);
                btn_Connect.Enabled = false;        // 查詢過程相當漫長，為了防止 User 手賤一直給我點連線按鈕，我乾脆直接把按鈕關掉直到查詢完畢。
                client.StartQuery();

            }
        }

        private bool ValidateConnect(QueryType qType)
        {
            bool v_server = true,v_port = true,v_filepath = true,v_folderpath = true,v_outfiletype = true;
            string[] ip = null;
            if (cb_ServerList.SelectedItem != null)
            {
                serverList = new ServerList { Value = cb_ServerList.SelectedItem.ToString() };
                ip = cb_ServerList.SelectedItem.ToString().Split('.');
            }
            else
            {
                if (!string.IsNullOrEmpty(cb_ServerList.Text))
                {
                    ip = cb_ServerList.Text.ToString().Split('.');
                    serverList = new ServerList { Value = cb_ServerList.Text };
                }                    
                else
                {
                    ip = "0.0.0.0".Split('.');
                    serverList = new ServerList { Value = "0.0.0.0" };
                }
            }
            int port;

            if (qType.Equals(QueryType.Server))
            {
                // 判斷Server IP是否填寫正確
                if (ip.Length == 4)
                {
                    try
                    {
                        /*
                         * 判斷 IP 這我沒有用正規表示式，因為要寫蠻長的(限制 0 ~ 255 的範圍)
                         * 相反的我是直接把字串用 "." 拆開後，針對拆開後的每個字串轉 int，並判斷有沒有在合法範圍之內
                         * 如果字串不能轉為int，跳出例外的話，也表示這值不正確。
                         */
                        v_server = ip.All(i => int.Parse(i) >= 0 && int.Parse(i) <= 255);
                    }
                    catch (Exception e)
                    {
                        v_server = false;
                    }
                }
                else
                {
                    v_server = false;
                }
                if (!v_server)
                    MessageBox.Show("Server端IP設定有誤，請檢查");

                // 判斷port是否填寫正確
                v_port = int.TryParse(txt_Port.Text, out port);
                if (!v_port || (port < 1025 || port > 65535))
                    MessageBox.Show("連接埠設定有誤，請檢查。(範圍：1025~65535)");
            }

            // 判斷檔案列表路徑是否正確
            txt_FilePath.Text = txt_FilePath.Text.Replace("\"", string.Empty);
            v_filepath = File.Exists(txt_FilePath.Text);
            if(!v_filepath || !txt_FilePath.Text.EndsWith(".xls"))
            {
                MessageBox.Show("檔案列表路徑設定有誤，請檢查。(僅接受.xls類型檔案)");
            }

            // 判斷儲存路徑是否正確
            txt_SaveFolder.Text = txt_SaveFolder.Text.Replace("\"", string.Empty);
            v_folderpath = Directory.Exists(txt_SaveFolder.Text);
            if (!v_folderpath)
                MessageBox.Show("儲存路徑設定有誤，請檢查。");

            if (!(v_outfiletype = (cb_OutputType.SelectedItem != null)))
                MessageBox.Show("請選擇檔案輸出類型。");
            

            return v_server && v_port && v_filepath && v_folderpath && v_outfiletype;
        }

        private void btn_Query_Click(object sender, EventArgs e)
        {
            if(ValidateConnect(QueryType.Local))
            {
                //QueryAction qAction = new QueryAction(txt_FilePath.Text, txt_SaveFolder.Text, this);
                btn_Query.Enabled = false;
                LocalQuery lQuery = new LocalQuery(txt_FilePath.Text, txt_SaveFolder.Text, this);
                lQuery.StartQuery();
                btn_Query.Enabled = true;
            }
        }
    }
}
