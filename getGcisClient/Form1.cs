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

        public Form1()
        {
            InitializeComponent();
            txt_ServerIP.Text = ConfigurationManager.AppSettings["serverIP"];
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
            dialog.Filter = "xls files (*.*)|*.xls";
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
            if(ValidateConnect())
            {
                Console.WriteLine("資料驗證成功...準備連線至 {0}，連接埠 {1}...",txt_ServerIP.Text,txt_Port.Text);

                Client client = new Client(txt_ServerIP.Text, txt_Port.Text, txt_FilePath.Text, txt_SaveFolder.Text, this);
                btn_Connect.Enabled = false;        // 查詢過程相當漫長，為了防止 User 手賤一直給我點連線按鈕，我乾脆直接把按鈕關掉直到查詢完畢。
                client.Connect();

            }
        }

        private bool ValidateConnect()
        {
            bool v_server = true,v_port = true,v_filepath = true,v_folderpath = true,v_outfiletype = true;
            var ip = txt_ServerIP.Text.Split('.');
            int port;            

            // 判斷Server IP是否填寫正確
            if(ip.Length == 4)
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
                catch(Exception e)
                {
                    v_server = false;
                }                
            }
            else
            {
                v_server = false;
            }
            if(!v_server)
                MessageBox.Show("Server端IP設定有誤，請檢查");

            // 判斷port是否填寫正確
            v_port = int.TryParse(txt_Port.Text, out port);
            if (!v_port || (port < 1025 || port > 65535))
                MessageBox.Show("連接埠設定有誤，請檢查。(範圍：1025~65535)");

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

        
    }
}
