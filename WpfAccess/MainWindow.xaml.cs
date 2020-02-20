using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Threading;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Xml;

namespace WpfAccess
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.DoWork += DoWork_Handler;
            bgWorker.ProgressChanged += ProgressChanged_Handler;
            bgWorker.RunWorkerCompleted += RunWorkerCompleted_Handler;
        }
        BackgroundWorker bgWorker = new BackgroundWorker();
        private string connectionString;
        private OleDbConnection odcConnection;
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\129.202.101.220\new\test.accdb;Jet OLEDB:database password=jay123456;Persist Security Info=False";
            try
            {
                // 建立连接  
                this.odcConnection = new OleDbConnection(this.connectionString);


                // 打开连接  
                this.odcConnection.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            intest.Visibility = Visibility.Hidden;
            /*  try
              {
                  this.odcConnection.Close();
              }
              catch (System.Exception ex)
              {
                  MessageBox.Show(ex.ToString());

              }
              */


        }

        private void Test_Click(object sender, RoutedEventArgs e)
        {
            if (!bgWorker.IsBusy)
            {
                bgWorker.RunWorkerAsync();
            }
            /* string Insert = "INSERT INTO S7TEST(File) values('" + "ED0"  + "')";

             //insert into 表名(字段1，字段2...)values('字段一内容'，'字段二内容')，上一行+用于字符串的连接，如果想用textBox传值，可用

             //string s = "'" + textBox1.Text + "'", x = "'" + textBox2.Text + "'";
             try
             {
                 OleDbCommand myCommand = new OleDbCommand(Insert, odcConnection);//执行命令

                 myCommand.ExecuteNonQuery();//更新数据库，返回受影响行数;可通过判断其是否>0来判断操作是否成功
             }
             catch (System.Exception ex)
             {
                 MessageBox.Show(ex.ToString());

             }
             */

        }

        private void DoWork_Handler(object sender, DoWorkEventArgs args)
        {
            string[] sn = new string[100];
            string[] snstatus = new string[100];
            string date = "";
          
            DirectoryInfo TheFolder = new DirectoryInfo("E:\\ERSA\\Export");
            do
            {
                foreach (FileInfo NextFile in TheFolder.GetFiles())
                {
                    DataSet ds = new DataSet();
                    DataTable dt = new DataTable();
                    DataSet ds2 = new DataSet();
                    DataTable dt2 = new DataTable();
                    if (NextFile.Name != "")
                    {
                      //  MessageBox.Show(NextFile.Name);
                        string Select = "SELECT *FROM S7TEST WHERE File = '" + NextFile.Name+"'";
                        //临时存储
                        OleDbDataAdapter inst=new OleDbDataAdapter();
                        try
                        {
                             inst = new OleDbDataAdapter(Select, odcConnection);//
                        }
                        catch (System.Exception ex)
                        {
                           MessageBox.Show(ex.ToString());

                        }
                       // MessageBox.Show("inst ok");
                        try
                        {
                            inst.Fill(ds,"Filename");//用inst填充ds 
                          
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.ToString());

                        }
                       //  MessageBox.Show("fill ok");
                         dt = ds.Tables["Filename"];
                      //    MessageBox.Show(dt.Rows.Count.ToString());
                        if (dt.Rows.Count == 0)
                        {
                            int pcbcount=0;
                            string Insert = "INSERT INTO S7TEST(File) values('" + NextFile.Name + "')";
                            try
                            {
                                OleDbCommand myCommand = new OleDbCommand(Insert, odcConnection);//执行命令

                                myCommand.ExecuteNonQuery();//更新数据库，返回受影响行数;可通过判断其是否>0来判断操作是否成功
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.ToString());

                            }

                            XmlTextReader reader = new XmlTextReader(NextFile.FullName);
                            while (reader.Read())
                            {
                                //if (reader.Name == "serial_pcb_1")
                                if (reader.Name.IndexOf("pcbs_in_panel") != -1)
                                { pcbcount = int.Parse(reader.ReadElementString().Trim()); }
                                    if (reader.Name.IndexOf("serial_pcb")!=-1)
                                {
                                    int index = int.Parse(reader.Name.Substring(reader.Name.Length-1,1));
                                   // if (reader.ReadElementString().Trim() != "")
                                     sn[index - 1] = reader.ReadElementString().Trim(); 
                                   // else
                                   // { sn[index - 1] = "0000"; }
                                }
                                if (reader.Name.IndexOf("status_pcb") != -1)
                                {
                                    int index = int.Parse(reader.Name.Substring(reader.Name.Length - 1, 1));
                                    snstatus[index - 1] = reader.ReadElementString().Trim();
                                }
                                if (reader.Name == "date")
                                { date= reader.ReadElementString().Trim(); }
                            }//reader.Read()
                          //  MessageBox.Show("Read end");
                            for (int i=0;i<pcbcount;i++)
                            {
                                if (sn[i] != "")
                                {
                                    string Select2 = "SELECT *FROM S7TEST2 WHERE SN = '" + sn[i] + "'";
                                   // MessageBox.Show(Select2);
                                    OleDbDataAdapter inst2 = new OleDbDataAdapter();
                                    try
                                    {
                                        inst2 = new OleDbDataAdapter(Select2, odcConnection);//
                                    }
                                    catch (System.Exception ex)
                                    {
                                        MessageBox.Show(ex.ToString());

                                    }
                                    try
                                    {
                                        inst2.Fill(ds2, "ProductionStatus");//用inst填充ds 

                                    }
                                    catch (System.Exception ex)
                                    {
                                        MessageBox.Show(ex.ToString());

                                    }
                                    dt2 = ds2.Tables["ProductionStatus"];
                                    if (dt2.Rows.Count == 0)
                                    {
                                        string Insert2 = "INSERT INTO [S7TEST2]([Filename],[SN],[Result],[Time]) values('" + NextFile.Name + "'," + "'" + sn[i] + "'," + "'" + snstatus[i] + "'," + "'" + date + "')";
                                        //  MessageBox.Show(Insert2);
                                        try
                                        {
                                            OleDbCommand myCommand2 = new OleDbCommand(Insert2, odcConnection);//执行命令

                                            myCommand2.ExecuteNonQuery();//更新数据库，返回受影响行数;可通过判断其是否>0来判断操作是否成功
                                        }
                                        catch (System.Exception ex)
                                        {
                                            MessageBox.Show(ex.ToString());

                                        }
                                    }
                                    if (dt2.Rows.Count != 0)
                                    {
                                        if ((ds2.Tables[0].Rows[0][3].ToString() == "NG") && (snstatus[i] == "OK"))
                                        {
                                            string Update = "UPDATE S7TEST2 SET Result='OK' WHERE SN='" + sn[i] + "'";
                                            try
                                            {
                                                OleDbCommand myCommand3 = new OleDbCommand(Update, odcConnection);//执行命令

                                                myCommand3.ExecuteNonQuery();//更新数据库，返回受影响行数;可通过判断其是否>0来判断操作是否成功
                                            }
                                            catch (System.Exception ex)
                                            {
                                                MessageBox.Show(ex.ToString());

                                            }
                                        }

                                    }

                                }//sn[i]!="0000"


                                        }//for(int i=0;i<pcbcount;i++)

                        }//dt.Rows.Count == 0
                       






                    }//ed0




                }//foreach

                System.Threading.Thread.Sleep(500);
            } while (true);



        }
        private void ProgressChanged_Handler(object sender, ProgressChangedEventArgs args)
        {

        }
        private void RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs args)
        {

        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                this.odcConnection.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private void In_Click(object sender, RoutedEventArgs e)
        {
            
            XmlTextReader reader = new XmlTextReader("D:\\1.xml");
            while (reader.Read())
            {
                //if (reader.Name == "serial_pcb_1")
                if (reader.Name.IndexOf("serial_pcb") != -1)
                {
                    string sbuf = reader.Name;
                    //  MessageBox.Show(reader.Name);
                    // if (reader.ReadElementString().Trim() == null)
                    //{ MessageBox.Show("Null"); }
                    if (reader.ReadElementString().Trim() == "")
                    { MessageBox.Show(sbuf); }

                }
                              
            }//reader.Read()

        }






    }
}
