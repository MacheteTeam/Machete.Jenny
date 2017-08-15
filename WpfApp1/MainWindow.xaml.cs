using Machete.Jenny;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly string SaveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);//默认在文档文件夹中

        private string SaveFileName = @"DXY";

        private string StrHeaderText = @"用户测试";

        Dictionary<string, string> titleNames = new Dictionary<string, string>();

        public MainWindow()
        {
            InitializeComponent();

            titleNames.Add("Name", "姓名");
            titleNames.Add("Age", "年龄");

            User user1 = new User();
            User user2 = new User();
            User user3 = new User();

            user1.Name = "Alex";
            user1.Age = 24;
            user1.Phone = "188";
            user2.Name = "Jenny";
            user2.Age = 22;
            user2.Phone = "4529";
            user3.Name = "yin";
            user3.Age = 0;
            user3.Phone = "6347";

            List<User> userList = new List<User>();
            userList.Add(user1);
            userList.Add(user2);
            userList.Add(user3);

            ExcelHelper.ListToExcel(userList, null, true, SaveFilePath, StrHeaderText, titleNames);
        }

        public class User
        {
            public string Name { get; set; }

            public int Age { get; set; }

            public string Phone { get; set; }
        }

        /// <summary>
        /// 选择Excel文件导入DataGridView
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChooseExcelFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = @"Excel文件|*.xls;*.xlsx" };  //;*.xlsx
            if (ofd.ShowDialog() == true)
            {
                try
                {

                    Type type = typeof(User);

                    //List<User> tList = ExcelHelper.ExcelToList<User>(type, ofd.FileName,0 ,true, titleNames)

                }
                catch (Exception ex)
                {
                    MessageBox.Show(@"导入失败！");
                }
            }
        }
    }
}
