using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Threading;

namespace WebApplication1
{
    public partial class _Default : System.Web.UI.Page
    {
        public class MachineInformation     //機台下詳細資訊的class(紀錄今天資料)
        {
            public string Name;             //機台編號
            public int Index;               //機台陣列所引
            public string StartTime;        //機台開始運轉時間
            public string EndTime;          //機台結束轉時間
            public int TotalWorkTimes;      //機台總運作時間
            public int TotalProduct;        //機台總產量
            public float TotalAvearage;    //機台平均產量
        }
        public class old_MachineInformation     //機台下詳細資訊的class(紀錄昨天資料)
        {
            public string Name;             //機台編號
            public int Index;               //機台陣列所引
            public string StartTime;        //機台開始運轉時間
            public string EndTime;          //機台結束轉時間
            public int TotalWorkTimes;      //機台總運作時間
            public int TotalProduct;        //機台總產量
            public float TotalAvearage;    //機台平均產量
        }

        

        //下載檔名
        //HttpCookie DateFileName_2 = new HttpCookie("");
        //private string[] DateFileName = { "", "" };

        //分析檔名
        //private string[] AnalysisFileName = { "", "" };

        //紀錄檔案下載時間
        //private DateTime RecodeDownloadDate;

        //紀錄檔案分析時間
        public DateTime RecodeAnalysisDate;

        //紀錄是否有按下DownLoad鍵
        //private bool isDownLoaded = true;

        //目前資料高度固定1440(扣除機台名稱等資訊)
        //private int DataHeight = 1440;
        private int DataHeight; //(已可自動找尋資料高度)

        //資料寬度會自行判斷
        private int DataWidth = 0;

        //存放所有資料陣列
        private string[][] ALLData = null;  //(紀錄當天資料)

        private string[][] ALLData_old = null; //(紀錄前一天資料)

        private int[] NameData = null;  //(紀錄Year記數資料 兩天分別計算 可用同一個變數)


        //存放每一個機台下的詳細資料
        public MachineInformation[] MachineDetail = null;  //(紀錄當天資料)
        public old_MachineInformation[] old_MachineDetail = null; //(紀錄前一天資料)

        string[] strDateYear = { "", "" }; //放置年(昨天 今天)
        string[] strDateMonth = { "", "" };//放置月(昨天 今天)
        string[] strDateDay = { "", "" };//放置日(昨天 今天)

        string MachineList_SAVE;  //機台名稱暫存空間(尚未拆解字串)


        public static string FTPAddress;
        public static string UserName;
        public static string PassWord;
        public static string SaveLocation;
        public static string Terminal;


        protected void Page_Load(object sender, EventArgs e)
        {
            Session.Add("Error",0);
            CheckConfig();
            //Response.Close();
        }

        private void CheckConfig()
        {
            //讀取設定檔
            /*if (File.Exists(Application.StartupPath + @"\config.boltun"))
            {
                System.IO.StreamReader file = new System.IO.StreamReader(Application.StartupPath + @"\config.boltun");
                FTPAddress = file.ReadLine();
                UserName = file.ReadLine();
                PassWord = file.ReadLine();
                SaveLocation = file.ReadLine();
                Terminal = file.ReadLine();

                MachineList_SAVE = file.ReadLine(); //將讀取資料放至機台名稱暫存空間(尚未拆解字串)

                /*txtAddress.Text = FTPAddress;
                txtUserName.Text = UserName;
                txtPassWord.Text = PassWord;
                txtLocation.Text = SaveLocation;
                txtTerminal.Text = Terminal;*/
               /* file.Close();
            }
            else//讀不到設定檔時顯示預設資料*/
            {
                FTPAddress = "ftp://10.1.12.100";
                UserName = "KVIE";
                PassWord = "1234";
                SaveLocation = "C:\\生產資料_A廠";
                Terminal = "/MMC/OOATWTN1A";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,11B06,14B19,11B01,11B09,11B02,11B10,14B03,14B09,14B08,14B12,11B14,14B22,11B03,NF120,14B11,14B20,14B16,14B15,11B05,11B04,14B04,14B21,17B02,17B01,14B24,19B34,19B33,19B32,19B31,19B30,19B29";
                

                /*FTPAddress = "ftp://10.1.12.171";
                UserName = "KVIE";
                PassWord = "1234";
                //SaveLocation = "C:\\生產資料";
                SaveLocation = "C:\\生產資料_B廠";
                Terminal = "/MMC/OOATWTN6A12";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,08B01,27B06,19B05,19B06,NF2412,27B04,27B01,14B06,14B23,14B17,14B18,14B13,14B07,11B07,14B05";
                */
            }
        }


        protected void dateTimeAnalysis_SelectionChanged(object sender, EventArgs e)
        {
            DateTime ChooseDateAnalysis = dateTimeAnalysis.SelectedDate;
            string[] DateFileName = { "", "" };
            
            if (DateTime.Compare(ChooseDateAnalysis, DateTime.Now.AddDays(0)) > 0)
            {
                Response.Write("<Script language='JavaScript'>alert('不能選取明天以後的日期！');</Script>");

                dateTimeAnalysis.SelectedDate = RecodeAnalysisDate;
                return;
            }
            //今天
            //將所選擇的日期轉換成檔名
            strDateYear[1] = ChooseDateAnalysis.Year.ToString();

            //切割字串:2015->15
            strDateYear[1] = strDateYear[1].Substring(2, 2);
            strDateMonth[1] = ChooseDateAnalysis.Month.ToString();
            strDateDay[1] = ChooseDateAnalysis.Day.ToString();
            //AnalysisFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";
            DateFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";
           
            //昨天
            //將所選擇的日期轉換成檔名
            strDateYear[0] = ChooseDateAnalysis.AddDays(-1).Year.ToString();
            //切割字串:2015->15
            strDateYear[0] = strDateYear[0].Substring(2, 2);
            strDateMonth[0] = ChooseDateAnalysis.AddDays(-1).Month.ToString();
            strDateDay[0] = ChooseDateAnalysis.AddDays(-1).Day.ToString();

            //AnalysisFileName[0] = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";
            DateFileName[0] = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";

            Session.Add("DateFileName[0]", DateFileName[0]);
            Session.Add("DateFileName[1]", DateFileName[1]);

            //Session.Add("AnalysisFileName[0]", AnalysisFileName[0]);
            //Session.Add("AnalysisFileName[1]", AnalysisFileName[1]);


            LbDetail.Visible = false;
            GridView_Detail.Visible = false;

            //清空所有資料
            GridView_Analysis.DataBind();

            /*string Error;
            try
            {
                Error = Session["Error"].ToString();
            }
            catch (Exception ex)
            {
                Session.Add("Error", 0);
                Error = Session["Error"].ToString();
                Response.Write(ex.ToString());
            }*/


            for (int i = 0; i < 2; i++)
            {
                if (Int32.Parse(Session["Error"].ToString()) == 0)//只要其中一天讀取錯誤則直接離開(不會跳出錯誤兩次)
                {
                    if (File.Exists(SaveLocation + "\\" + DateFileName[i]))
                    {

                        //lbDate.Text = "分析檔案 : " + DateFileName[i];
                        //LbDetail.Text = "";
                        //取得機器名稱，並把資料放入陣列
                        GreateMachineData(SaveLocation + "\\" + DateFileName[i], i);
                    }
                    else
                    {
                        Session.Add("Error", 1);
                        //Error = 1;
                        string[] DayList = DateFileName[i].Split('.');
                        if (i == 0)
                        {
                            Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容，並重新下載！');</Script>");
                        }
                        else if (i == 1)
                        {
                            Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容，並重新下載！');</Script>");
                        }
                        break;
                    }

                }

                else if (Int32.Parse(Session["Error"].ToString()) == 1)//只要其中一天讀取錯誤則直接離開(不會跳出錯誤兩次)
                {
                    break;
                }
            }
        }


        protected void Btn_Download_Click(object sender, EventArgs e)
        {
            string[] DateFileName = { "", "" };
            DateFileName[0] = Session["DateFileName[0]"].ToString();
            DateFileName[1] = Session["DateFileName[1]"].ToString(); 

            try
            {
                if (SaveLocation == "")
                {
                    Response.Write("<Script language='JavaScript'>alert('請輸入存放位置！');</Script>");
                    return;
                }
                if (FTPAddress == "" || UserName == "" || PassWord == "")
                {
                    Response.Write("<Script language='JavaScript'>alert('請輸入完整資料！');</Script>");
                    return;
                }
                
                //下載兩天資料
                for (int i = 0; i < 2; i++)
                {
                    ftp ftpclient = new ftp(FTPAddress, UserName, PassWord);
                    string Location = Terminal + "/" + DateFileName[i];
                    Console.WriteLine(Location);

                    int isSuccess = DownloadWithProgressBar(FTPAddress ,UserName, PassWord , Terminal + "/" + DateFileName[i], SaveLocation + "\\" + DateFileName[i]);

                    if (isSuccess == -1)
                    {
                        Response.Write("<Script language='JavaScript'>alert('下載檔案失敗，請檢查檔案是否存在！');</Script>");
                        ftpclient = null;
                        break;
                    }
                    else if (isSuccess == -2)
                    {
                        Response.Write("<Script language='JavaScript'>alert('登入失敗，請檢查輸入帳號密碼是否正確！');</Script>");
                        break;
                    }
                    if (i == 1)
                    {
                        Session.Add("Error", 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
            }
        }

        
        //int DataHeight_yesterday; //計算昨天資料數量(可防止昨天與今天資料量不同也可進行分析)
        //int DataHeight_today;   //計算今天資料數量(可防止昨天與今天資料量不同也可進行分析)

        //計算每一機台的整天統計資料
        private void GreateMachineData(string AnalysisFileName, int input_day)  //input_day = 0為昨天資料、input_day = 1為今天資料
        {
            DataTable dt = new DataTable();
            DataRow dr = null;

            //確定可以寫入
            dt.Columns.Add(new DataColumn("機台名稱", typeof(string)));
            dt.Columns.Add(new DataColumn("起始時間", typeof(string)));
            dt.Columns.Add(new DataColumn("結束時間", typeof(string)));
            dt.Columns.Add(new DataColumn("總產量", typeof(string)));
            dt.Columns.Add(new DataColumn("總運行時間", typeof(string)));
            dt.Columns.Add(new DataColumn("平均產能", typeof(string)));
            
                
            //測試View貼值  
            /*dt.Columns.Add(new DataColumn("RowNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("Column1", typeof(string)));
            dt.Columns.Add(new DataColumn("Column2", typeof(string)));
            dt.Columns.Add(new DataColumn("Column3", typeof(string)));

            for (int i = 0; i < 100; i++)
            {
                dr = dt.NewRow();
                dr["RowNumber"] = i;
                dr["Column1"] = i + 100;
                dr["Column2"] = i + 1000;
                dr["Column3"] = i + 10000;
                dt.Rows.Add(dr);
            }

            //Store the DataTable in ViewState
            ViewState["CurrentTable"] = dt;

            GridView_Analysis.DataSource = dt;
            GridView_Analysis.DataBind();*/
            //測試View貼值


            try
            {
                //目前資料所引
                int ProcessCounter = 0;

                //每行字串
                string EachLine = "";

                //讀取機器名稱
                string MachineData = "";

                int DeleteValue = 0;    //自動判斷不要資料的索引位置在哪裡的計數值

                int DeleteValue_Num = 0;//判斷是否有多行資料要刪除

                int DeleteValue_Count = 0;//判斷多行要刪除資料計數 如果到達多行刪除資料量即要停止

                string[] MachineList; //顯示機台名稱用
                //讀取檔案
                System.IO.StreamReader file = new System.IO.StreamReader(AnalysisFileName);

                //自行判斷資料數量使用
                int ReadLine = 0;

                DataHeight = 2000;

                for (int i = 0; i < DataHeight; i++)
                {
                    MachineData = file.ReadLine();

                    if (MachineData == null) //如果讀取不到資料直接跳出迴圈
                    {
                        file.Close(); //須關閉資料(否則下次會讀取不到資料)
                        file = new System.IO.StreamReader(AnalysisFileName);//讀取檔案
                        break;
                    }

                    //判斷是否有多行資料要刪除
                    if (MachineData.IndexOf("Year") != -1)
                    {
                        DeleteValue_Num++;//判斷要進行幾行刪除資料(累加出來的結果為刪除數量)
                    }

                    ReadLine++;//只要資料還讀取的到則一直累加
                }

                DataHeight = ReadLine; //將累加的資料放至要跑得迴圈內

                NameData = new int[DeleteValue_Num + 1];
                int NameData_num = 0;
                for (int i = 0; i < DataHeight; i++)
                {
                    MachineData = file.ReadLine();
                    
                    if (MachineData.IndexOf("Year") != -1)
                    {
                        //if (input_day == 0)
                        {
                            NameData[NameData_num] = i;
                            NameData_num++;
                        }
                        /*else if (input_day == 1)
                        {
                            NameData[NameData_old_num] = i;
                            NameData_num++;
                        }*/

                    }

                }

                file.Close(); //須關閉資料(否則下次會讀取不到資料)
                file = new System.IO.StreamReader(AnalysisFileName);//讀取檔案

                //只要找到"Year"字眼則跳出(抓出第一行)
                for (int i = 0; i < DataHeight; i++)
                {
                    MachineData = file.ReadLine();
                    if (MachineData.IndexOf("Year") != -1)
                    {
                        /*if (DeleteValue >= DataHeight)  //如果找不到刪除不要資料的索引則以第一行當成設定值
                            DeleteValue = 1;*/
                        break;
                    }
                    DeleteValue++;//自動判斷不要資料的索引位置在哪裡的計數值累加
                }  
                

                if (DeleteValue < DataHeight)//如果有抓到機台資料
                {
                    MachineList_SAVE = MachineData; //將機台資料先丟至暫存記憶體等待程式關閉時儲存或是等下一天抓不到資料時可以有資料參考
                }

                //空間宣告
                if (input_day == 0)
                {
                    ALLData_old = new string[DataHeight][];
                }
                else if (input_day == 1)
                {
                    ALLData = new string[DataHeight][];
                }

                if (DeleteValue < DataHeight) //如果有找到機台名稱
                {
                    //字串分割出每台機器名稱，並加入至下拉式選單(有加value tag)
                    MachineList = MachineData.Split(',');
                }
                else//如果找尋不到機台名稱
                {
                    //將暫存記憶體 字串分割出每台機器名稱 丟至List(如果兩天都沒有機台資料 也會有資料可以參考)
                    MachineList = MachineList_SAVE.Split(',');
                }

                DataWidth = MachineList.Length;
                // 2015.09.25彥昌 新增 如果MachineList最後一個值為空 則不讀取最後一個值(修改掉之前強制把最後一個逗號拿掉會造成機台名稱少一個數字)
                for (int i = 0; i < DataWidth; i++)
                {
                    if (MachineList[i] == "")
                    {
                        DataWidth = i;
                        break;
                    }
                }

                if (input_day == 0)//(前一天資料)
                {
                    old_MachineDetail = new old_MachineInformation[DataWidth - 6];
                    for (int i = 6; i < DataWidth; i++)
                    {
                        old_MachineDetail[i - 6] = new old_MachineInformation();
                        old_MachineDetail[i - 6].Name = MachineList[i];
                        old_MachineDetail[i - 6].Index = i;
                    }
                }
                else if (input_day == 1)//(當天資料)
                {
                    MachineDetail = new MachineInformation[DataWidth - 6];
                    for (int i = 6; i < DataWidth; i++)
                    {
                        MachineDetail[i - 6] = new MachineInformation();
                        MachineDetail[i - 6].Name = MachineList[i];
                        MachineDetail[i - 6].Index = i;
                    }
                }

                //關閉檔案並重新讀取檔案
                file.Close();
                file = new System.IO.StreamReader(AnalysisFileName);

                ReadLine = 0;//消除掉不需要判斷資料後 要將計數清空(重新計數)
                //把資料放入陣列裡
                for (int i = 0; i < DataHeight; i++)
                {
                    //如果計數到了要跳過那行 和 如果有多行要刪除資料而且計數尚未超過要刪除的資料量
                    if (ProcessCounter == DeleteValue && DeleteValue_Num > DeleteValue_Count)
                    {

                        DeleteValue_Count++;
                        //DeleteValue += 61; //如果只有一行機台名稱 即可刪除此行(Year,month,day,hour,minute)
                        DeleteValue = NameData[DeleteValue_Count];

                        EachLine = file.ReadLine();
                        ProcessCounter++;
                    }
                    EachLine = file.ReadLine();
                    if (EachLine == null)//當讀取不到值時則跳出
                    {
                        break;
                    }
                    ReadLine++;//只要有讀取到資料則累計計數
                    //二維陣列的宣告
                    if (input_day == 0)//(前一天資料)
                    {
                        ALLData_old[i] = new string[DataWidth];
                    }
                    else if (input_day == 1)//(當天資料)
                    {
                        ALLData[i] = new string[DataWidth];
                    }
                    //string eee;
                    for (int j = 0; j < DataWidth; j++)
                    {
                        string[] SplitEachLine = EachLine.Split(',');

                        if (input_day == 0)//(前一天資料)
                        {
                            ALLData_old[i][j] = SplitEachLine[j];

                            Session.Add("ALLData_old[" + i.ToString() + "][" + j.ToString() + "]".ToString(), SplitEachLine[j]);
                        }
                        else if (input_day == 1)//(當天資料)
                        {
                            ALLData[i][j] = SplitEachLine[j];

                            Session.Add("ALLData[" + i.ToString() + "][" + j.ToString() + "]".ToString(), SplitEachLine[j]);
                        }
                    }
                    ProcessCounter++;
                }

                DataHeight = ReadLine;//跳出後將計數資料放至須跑之資料內
                if (input_day == 0)//判斷為昨天資料
                {
                    //DataHeight_yesterday = ReadLine;//跳出後將計數資料放至昨天資料內
                    Session.Add("DataHeight_yesterday", ReadLine);//計算昨天資料數量(可防止昨天與今天資料量不同也可進行分析)  
                }
                else if (input_day == 1)//判斷為今天資料
                {
                    //DataHeight_today = ReadLine;//跳出後將計數資料放至今天資料內
                    Session.Add("DataHeight_today", ReadLine);//計算今天資料數量(可防止昨天與今天資料量不同也可進行分析)
                }
                //關閉檔案
                file.Close();
                
                //計算每台機器所有資料
                for (int i = 0; i < DataWidth - 6; i++)
                {
                    if (input_day == 0)//(昨天資料分析)
                    {
                        string StartTime = ALLData_old[1][3].PadLeft(2, '0') + " : " + ALLData_old[1][4].PadLeft(2, '0');
                        string EndTime = ALLData_old[DataHeight - 1][3].PadLeft(2, '0') + " : " + ALLData_old[DataHeight - 1][4].PadLeft(2, '0');
                        old_MachineDetail[i].StartTime = StartTime;
                        old_MachineDetail[i].EndTime = EndTime;
                        old_MachineDetail[i].TotalProduct = 0;
                        old_MachineDetail[i].TotalWorkTimes = 0;
                        old_MachineDetail[i].TotalAvearage = 0;
                        for (int j = 1; j < DataHeight - 1; j++)
                        {
                            //取得當前產量與下一分鐘的產量
                            int ThisValue = Convert.ToInt32(ALLData_old[j][old_MachineDetail[i].Index]);
                            int NextValue = Convert.ToInt32(ALLData_old[j + 1][old_MachineDetail[i].Index]);
                            //每分鐘產量>10顆
                            if (NextValue - ThisValue > 10)
                            {
                                //總產能與總運作時間
                                old_MachineDetail[i].TotalProduct += NextValue - ThisValue;
                                old_MachineDetail[i].TotalWorkTimes++;
                            }
                        }
                        
                        //平均產能
                        if (old_MachineDetail[i].TotalWorkTimes != 0)
                            old_MachineDetail[i].TotalAvearage = ((float)(old_MachineDetail[i].TotalProduct) / (float)(old_MachineDetail[i].TotalWorkTimes));
                        else
                            old_MachineDetail[i].TotalAvearage = 0;
                    }
                    //機台開始與結束時間
                    else if (input_day == 1)//(今天資料分析)
                    {
                        string StartTime = ALLData[1][3].PadLeft(2, '0') + " : " + ALLData[1][4].PadLeft(2, '0');
                        string EndTime = ALLData[DataHeight - 1][3].PadLeft(2, '0') + " : " + ALLData[DataHeight - 1][4].PadLeft(2, '0');
                        MachineDetail[i].StartTime = StartTime;
                        MachineDetail[i].EndTime = EndTime;
                        MachineDetail[i].TotalProduct = 0;
                        MachineDetail[i].TotalWorkTimes = 0;
                        MachineDetail[i].TotalAvearage = 0;
                        for (int j = 1; j < DataHeight - 1; j++)
                        {
                            //取得當前產量與下一分鐘的產量
                            int ThisValue = Convert.ToInt32(ALLData[j][MachineDetail[i].Index]);
                            int NextValue = Convert.ToInt32(ALLData[j + 1][MachineDetail[i].Index]);
                            //每分鐘產量>10顆
                            if (NextValue - ThisValue > 10)
                            {
                                //總產能與總運作時間
                                MachineDetail[i].TotalProduct += NextValue - ThisValue;
                                MachineDetail[i].TotalWorkTimes++;
                            }
                        }
                        //平均產能
                        if (MachineDetail[i].TotalWorkTimes != 0) MachineDetail[i].TotalAvearage = ((float)(MachineDetail[i].TotalProduct) / (float)(MachineDetail[i].TotalWorkTimes));
                        else MachineDetail[i].TotalAvearage = 0;
                    }

                    //新增至DataGridView
                    if (input_day == 1)//(當天資料 判斷時須拿前一天資料做整合)
                    {
                        //object[] obj;

                        if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
                        {
                            //obj = new object[6] { MachineDetail[i].Name, MachineDetail[i].StartTime, MachineDetail[i].EndTime, MachineDetail[i].TotalProduct, MachineDetail[i].TotalWorkTimes, Math.Round(MachineDetail[i].TotalAvearage) };
                            dr = dt.NewRow();
                            dr["機台名稱"] = MachineDetail[i].Name;
                            dr["起始時間"] = MachineDetail[i].StartTime;
                            dr["結束時間"] = MachineDetail[i].EndTime;
                            dr["總產量"] = MachineDetail[i].TotalProduct;
                            dr["總運行時間"] = MachineDetail[i].TotalWorkTimes;
                            dr["平均產能"] = Math.Round(MachineDetail[i].TotalAvearage);
                            //dr["Column3"] = i + 10000;
                            dt.Rows.Add(dr);
                        }
                        else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天")//只顯示昨日資料
                        {
                            //obj = new object[6] { old_MachineDetail[i].Name, old_MachineDetail[i].StartTime, old_MachineDetail[i].EndTime, old_MachineDetail[i].TotalProduct, old_MachineDetail[i].TotalWorkTimes, Math.Round(old_MachineDetail[i].TotalAvearage) };
                            dr = dt.NewRow();
                            dr["機台名稱"] = old_MachineDetail[i].Name;
                            dr["起始時間"] = old_MachineDetail[i].StartTime;
                            dr["結束時間"] = old_MachineDetail[i].EndTime;
                            dr["總產量"] = old_MachineDetail[i].TotalProduct;
                            dr["總運行時間"] = old_MachineDetail[i].TotalWorkTimes;
                            dr["平均產能"] = Math.Round(old_MachineDetail[i].TotalAvearage);
                            //dr["Column3"] = i + 10000;
                            dt.Rows.Add(dr);
                        }
                        else //顯示兩天(今、昨天)資料
                        {
                            //obj = new object[6] { MachineDetail[i].Name, old_MachineDetail[i].StartTime, MachineDetail[i].EndTime, (old_MachineDetail[i].TotalProduct + MachineDetail[i].TotalProduct), (old_MachineDetail[i].TotalWorkTimes + MachineDetail[i].TotalWorkTimes), Math.Round((old_MachineDetail[i].TotalAvearage + MachineDetail[i].TotalAvearage) / 2, 0) };
                            dr = dt.NewRow();
                            dr["機台名稱"] = MachineDetail[i].Name;
                            dr["起始時間"] = old_MachineDetail[i].StartTime;
                            dr["結束時間"] = MachineDetail[i].EndTime;
                            dr["總產量"] = (old_MachineDetail[i].TotalProduct + MachineDetail[i].TotalProduct);
                            dr["總運行時間"] = (old_MachineDetail[i].TotalWorkTimes + MachineDetail[i].TotalWorkTimes);
                            dr["平均產能"] = Math.Round((old_MachineDetail[i].TotalAvearage + MachineDetail[i].TotalAvearage) / 2);
                            //dr["Column3"] = i + 10000;
                            dt.Rows.Add(dr);
                        }
                        
                    }
                }
                ViewState["CurrentTable"] = dt;

                GridView_Analysis.DataSource = dt;
                GridView_Analysis.DataBind();
            }
            catch (Exception ex)
            {
                if (ex.ToString().IndexOf("並未將物件參考設定為物件的執行個體") != -1)
                {
                    if (input_day == 0)
                    {
                        Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容！');</Script>");
                    }
                    if (input_day == 1)
                    {
                        Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容！');</Script>");
                    }
                }
                else
                {
                    if (input_day == 0)
                    {
                        Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容！');</Script>");
                    }
                    if (input_day == 1)
                    {
                        Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容！');</Script>");
                    }
                }
                Session.Add("Error", 1);
            }
        }

        protected void GridView_Analysis_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            e.CommandArgument.ToString();
            sender.ToString();

            int index = Convert.ToInt16(e.CommandArgument.ToString());
            GridViewRow ChoseRow = GridView_Analysis.Rows[index];


            int Number_Day = 0;
            try
            {
                GridView_Detail.Visible = true;
                LbDetail.Visible = true;
                LbDetail.Text = ChoseRow.Cells[1].Text + " 機台詳細產能";

                if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
                {
                    Number_Day = 1;
                }
                else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天") //只顯示今日資料
                {
                    Number_Day = 0;
                }
                else
                {
                    Number_Day = 2;
                }

                AnalysisProduction(index, ChoseRow.Cells[1].Text, Number_Day);

            }
            catch (Exception ex)
            {
                Console.WriteLine("錯誤行數：{0}毫秒", ex.ToString());
            }
        }

        int Error_Day;//確認是否有錯誤
        //計算所選擇機台的詳細資料
        private void AnalysisProduction(int MachineIndex, string MachineName, int input_Num)
        {

            DataTable dt = new DataTable();
            DataRow dr = null;
            dt.Columns.Add(new DataColumn("起始時間", typeof(string)));
            dt.Columns.Add(new DataColumn("結束時間", typeof(string)));
            dt.Columns.Add(new DataColumn("平均值", typeof(string)));

            try
            {
                Error_Day = 0;
                //計數換顏色用(分單數跟雙數)
                int GridViewCounter = 0;

                //紀錄起始及結束時間
                string StartTime = "";
                string EndTime = "";

                //紀錄停機時間
                string StopStartTime = "";
                string StopEndTime = "";
                int StopProduce = 0;
                int StopCounting = 0;
                //int Count_Num = 0;

                //分鐘統計
                int MinuteCounter = 0;

                //數量統計
                int ProduceCounter = 0;


                //產量統計
                int ProductionCapacity = 0;

                //取得所選的ComboBox Tag做為二為陣列的索引
                int ComboTag = MachineIndex + 6;
                //int ComboTag = 6;

                //每分鐘產量閥值
                int ProductThreshold;

                //分析機台名稱
                if (MachineName.IndexOf("NF") != -1 || MachineName.IndexOf("17") != -1 || MachineName.IndexOf("19") != -1)
                    ProductThreshold = 50;
                else ProductThreshold = 100;

                if (input_Num == 0 || input_Num == 2) //昨天資料處理
                {
                    //dataGridView_Analysis.Rows.Clear(); //清空資料表
                    GridViewCounter = 0;
                    object[] obj_yesterday = new object[] { "昨天資料" };


                    dr = dt.NewRow();
                    dr["起始時間"] = "昨天資料";
                    dr["結束時間"] = "";
                    dr["平均值"] = "";
                    dt.Rows.Add(dr);

                    GridViewCounter++;

                    DataHeight = Convert.ToInt32(Session["DataHeight_yesterday"].ToString());   //將昨天之計數回傳至DataHeight內(可防止昨天與今天資料量不同也可進行分析)
                    //計算產量
                    for (int i = 1; i < DataHeight - 1; i++)
                    {

                        int ThisValue = Convert.ToInt32(Session["ALLData_old[" + i + "][" + ComboTag + "]"].ToString());
                        int NextValue = Convert.ToInt32(Session["ALLData_old[" + (i + 1) + "][" + ComboTag + "]"].ToString());

                        //每分鐘產量>100顆
                        if (NextValue - ThisValue > ProductThreshold)
                        {
                            if (StopProduce == 1)
                            {
                                StopEndTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');

                                object[] obj = new object[3] { StopStartTime, StopEndTime, "停機" };

                                dr = dt.NewRow();
                                dr["起始時間"] = StopStartTime;
                                dr["結束時間"] = StopEndTime;
                                dr["平均值"] = "停機";
                                //dr["Column3"] = i + 10000;
                                dt.Rows.Add(dr);

                                GridViewCounter++;
                                StopProduce = 0;
                                StopCounting = 0;
                            }

                            if (MinuteCounter == 0)
                                StartTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0'); //紀錄運轉起始時間
                            ProduceCounter += NextValue - ThisValue;
                            MinuteCounter++;


                            //到最後皆大於10顆
                            if (i == DataHeight - 2 && MinuteCounter != 0)
                            {
                                EndTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0'); //紀錄運轉結束時間
                                ProductionCapacity = ProduceCounter / MinuteCounter;
                                object[] obj = new object[3] { StartTime, EndTime, ProductionCapacity.ToString() };

                                dr = dt.NewRow();
                                dr["起始時間"] = StartTime;
                                dr["結束時間"] = EndTime;
                                dr["平均值"] = ProductionCapacity.ToString();
                                dt.Rows.Add(dr);
                                GridViewCounter++;
                            }
                        }
                        else
                        {
                            if (ProduceCounter != 0 && MinuteCounter != 0)
                            {
                                EndTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');  //紀錄運轉結束時間
                                ProductionCapacity = ProduceCounter / MinuteCounter;
                                object[] obj = new object[3] { StartTime, EndTime, ProductionCapacity.ToString() };//新增結果

                                GridViewCounter++;
                                StopProduce = 1;

                                dr = dt.NewRow();
                                dr["起始時間"] = StartTime;
                                dr["結束時間"] = EndTime;
                                dr["平均值"] = ProductionCapacity.ToString();
                                dt.Rows.Add(dr);
                            }
                            else if (StopProduce == 0)
                            {
                                //開始使停機
                                StopProduce = 1;
                                StopStartTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                            }
                            else if (StopProduce == 1 && i == DataHeight - 2)
                            {
                                //持續到最後一分鐘都停機，直接顯示                                    
                                StopEndTime = Session["ALLData_old[" + (i + 1) + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + (i + 1) + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                                object[] obj = new object[3] { StopStartTime, StopEndTime, "停機" };

                                dr = dt.NewRow();
                                dr["起始時間"] = StopStartTime;
                                dr["結束時間"] = StopEndTime;
                                dr["平均值"] = "停機";
                                dt.Rows.Add(dr);

                                GridViewCounter++;
                                StopProduce = 0;
                            }
                            if (StopProduce == 1 && StopCounting == 0)
                            {
                                //開始使停機                                    
                                StopProduce = 1;
                                StopStartTime = Session["ALLData_old[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData_old[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                                StopCounting = 1;
                            }


                            //變數重置
                            ProduceCounter = MinuteCounter = 0;
                            EndTime = StartTime = "";
                        }
                    }
                }


                //計算今天日期 所有變數都要從來
                Error_Day = 1;

                //紀錄起始及結束時間
                StartTime = "";
                EndTime = "";

                //紀錄停機時間
                StopStartTime = "";
                StopEndTime = "";
                StopProduce = 0;
                StopCounting = 0;
                //int Count_Num = 0;

                //分鐘統計
                MinuteCounter = 0;

                //數量統計
                ProduceCounter = 0;


                //產量統計
                ProductionCapacity = 0;

                //取得所選的ComboBox Tag做為二為陣列的索引
                ComboTag = MachineIndex + 6;
                //int ComboTag = 6;

                //每分鐘產量閥值


                //分析機台名稱
                if (MachineName.IndexOf("NF") != -1 || MachineName.IndexOf("17") != -1 || MachineName.IndexOf("19") != -1)
                    ProductThreshold = 50;
                else ProductThreshold = 100;


                if (input_Num == 1 || input_Num == 2)//今天資料處理
                {
                    GridViewCounter = 0;
                    object[] obj_today = new object[] { "今天資料" };

                    dr = dt.NewRow();
                    dr["起始時間"] = "今天資料";
                    dr["結束時間"] = "";
                    dr["平均值"] = "";
                    dt.Rows.Add(dr);

                    GridViewCounter++;

                    DataHeight = Convert.ToInt32(Session["DataHeight_today"].ToString());   //將今天之計數回傳至DataHeight內(可防止昨天與今天資料量不同也可進行分析)


                    //計算產量
                    for (int i = 1; i < DataHeight - 1; i++)
                    {
                        //取得當前產量與下一分鐘的產量
                        int ThisValue = Convert.ToInt32(Session["ALLData[" + i + "][" + ComboTag + "]"].ToString());
                        int NextValue = Convert.ToInt32(Session["ALLData[" + (i + 1) + "][" + ComboTag + "]"].ToString());
                        //每分鐘產量>100顆
                        if (NextValue - ThisValue > ProductThreshold)
                        {
                            if (StopProduce == 1)
                            {
                                //結束停機並新增資料
                                StopEndTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                                object[] obj = new object[3] { StopStartTime, StopEndTime, "停機" };

                                dr = dt.NewRow();
                                dr["起始時間"] = StopStartTime;
                                dr["結束時間"] = StopEndTime;
                                dr["平均值"] = "停機";
                                dt.Rows.Add(dr);


                                GridViewCounter++;
                                StopProduce = 0;

                                StopCounting = 0;
                            }

                            //時間資料對齊
                            if (MinuteCounter == 0) StartTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');     //紀錄運轉起始時間
                            ProduceCounter += NextValue - ThisValue;
                            MinuteCounter++;

                            //到最後皆大於10顆
                            if (i == DataHeight - 2 && MinuteCounter != 0)
                            {
                                EndTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');      //紀錄運轉結束時間
                                ProductionCapacity = ProduceCounter / MinuteCounter;
                                object[] obj = new object[3] { StartTime, EndTime, ProductionCapacity.ToString() };

                                dr = dt.NewRow();
                                dr["起始時間"] = StartTime;
                                dr["結束時間"] = EndTime;
                                dr["平均值"] = ProductionCapacity.ToString();
                                dt.Rows.Add(dr);

                                GridViewCounter++;
                            }
                        }
                        else
                        {
                            if (ProduceCounter != 0 && MinuteCounter != 0)
                            {
                                EndTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');  //紀錄運轉結束時間
                                ProductionCapacity = ProduceCounter / MinuteCounter;
                                object[] obj = new object[3] { StartTime, EndTime, ProductionCapacity.ToString() };//新增結果

                                dr = dt.NewRow();
                                dr["起始時間"] = StartTime;
                                dr["結束時間"] = EndTime;
                                dr["平均值"] = ProductionCapacity.ToString();
                                dt.Rows.Add(dr);

                                GridViewCounter++;
                                StopProduce = 1;
                            }
                            else if (StopProduce == 0)
                            {
                                //開始使停機                                    
                                StopProduce = 1;
                                StopStartTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                            }
                            else if (StopProduce == 1 && i == DataHeight - 2)
                            {
                                //持續到最後一分鐘都停機，直接顯示
                                StopEndTime = Session["ALLData[" + (i + 1) + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + (i + 1) + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                                object[] obj = new object[3] { StopStartTime, StopEndTime, "停機" };

                                dr = dt.NewRow();
                                dr["起始時間"] = StopStartTime;
                                dr["結束時間"] = StopEndTime;
                                dr["平均值"] = "停機";
                                dt.Rows.Add(dr);

                                GridViewCounter++;
                                StopProduce = 0;
                            }
                            if (StopProduce == 1 && StopCounting == 0)
                            {
                                //開始使停機
                                StopProduce = 1;
                                StopStartTime = Session["ALLData[" + i + "][" + 3 + "]"].ToString().PadLeft(2, '0') + ":" + Session["ALLData[" + i + "][" + 4 + "]"].ToString().PadLeft(2, '0');
                                StopCounting = 1;
                            }

                            //變數重置
                            ProduceCounter = MinuteCounter = 0;
                            EndTime = StartTime = "";
                        }
                    }

                }

                //最後結束時將所有運算結果貼至GridView_Detail
                ViewState["CurrentTable"] = dt;

                GridView_Detail.DataSource = dt;
                GridView_Detail.DataBind();

            }
            catch (Exception ex)
            {
                if (Error_Day == 0)
                {
                    Response.Write("昨天檔案有問題，請確認昨日檔案內容" + "\r\n" + ex.ToString() + "讀取檔案內容時發生錯誤，請檢查檔案格式");
                }
                else if (Error_Day == 1)
                {
                    Response.Write("今天檔案有問題，請確認今日檔案內容" + "\r\n" + ex.ToString() + "讀取檔案內容時發生錯誤，請檢查檔案格式");
                }

                Session.Add("Error", 1);
            }
        }

        Thread thread;
        int Running = 1;
        protected void RadioBtn_Viewdetails_Init(object sender, EventArgs e)
        {
            try
            {
                if (Session["Thread_Day"].ToString() != "1")
                {
 
                }
            }
            catch 
            {
                string[] DateFileName = { "", "" };
                //今天
                //將所選擇的日期轉換成檔名
                strDateYear[1] = System.DateTime.Now.AddDays(0).Year.ToString();
                //切割字串:2015->15
                //System.DateTime.Now.Year
                strDateYear[1] = strDateYear[1].Substring(2, 2);

                strDateMonth[1] = System.DateTime.Now.AddDays(0).Month.ToString();
                strDateDay[1] = System.DateTime.Now.AddDays(0).Day.ToString();
                //AnalysisFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";
                DateFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";

                //昨天
                //將所選擇的日期轉換成檔名
                strDateYear[0] = System.DateTime.Now.AddDays(-1).Year.ToString();
                //切割字串:2015->15
                strDateYear[0] = strDateYear[0].Substring(2, 2);

                strDateMonth[0] = System.DateTime.Now.AddDays(-1).Month.ToString();
                strDateDay[0] = System.DateTime.Now.AddDays(-1).Day.ToString();

                //AnalysisFileName[0] = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";
                DateFileName[0] = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";

                Session.Add("DateFileName[0]", DateFileName[0]);
                Session.Add("DateFileName[1]", DateFileName[1]);

                //Session.Add("AnalysisFileName[0]", AnalysisFileName[0]);
                //Session.Add("AnalysisFileName[1]", AnalysisFileName[1]);


                thread = new Thread(moveAuto); //啟動Thread(需要自動下載檔案)
                thread.Start();
                Session.Add("Thread_Day", "1");
                
            }
        }
        /*int Number_Day_Chang = 0;
        int statement1;*/
        private void moveAuto()
        {
            
            while (Running == 1)
            {
                if (System.DateTime.Now.Hour == 0 && System.DateTime.Now.Minute == 0 && System.DateTime.Now.Second == 0)
                {
                    //下載FTP檔案
                    Download();
                }

                Thread.Sleep(1000);
            }
        }


        private void Download()
        {
            //DateTime ChooseDateAnalysis = dateTimeAnalysis.SelectedDate;
            string DateFileName_Thread;

            /*strDateYear[1] = System.DateTime.Now.AddDays(-1).Year.ToString();

            //切割字串:2015->15
            //System.DateTime.Now.Year
            strDateYear[1] = strDateYear[1].Substring(2, 2);

            strDateMonth[1] = System.DateTime.Now.AddDays(-1).Month.ToString();
            strDateDay[1] = System.DateTime.Now.AddDays(-1).Day.ToString();
            //AnalysisFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";
            DateFileName[1] = strDateYear[1] + "." + strDateMonth[1] + "." + strDateDay[1] + ".csv";*/



            //昨天
            //將所選擇的日期轉換成檔名

            strDateYear[0] = System.DateTime.Now.AddDays(-1).Year.ToString();
            //切割字串:2015->15
            strDateYear[0] = strDateYear[0].Substring(2, 2);

            strDateMonth[0] = System.DateTime.Now.AddDays(-1).Month.ToString();
            strDateDay[0] = System.DateTime.Now.AddDays(-1).Day.ToString();

            //AnalysisFileName[0] = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";
            DateFileName_Thread = strDateYear[0] + "." + strDateMonth[0] + "." + strDateDay[0] + ".csv";

            /*Session.Add("DateFileName[0]", DateFileName[0]);
            Session.Add("DateFileName[1]", DateFileName[1]);

            Session.Add("AnalysisFileName[0]", AnalysisFileName[0]);
            Session.Add("AnalysisFileName[1]", AnalysisFileName[1]);*/

            /*DateFileName[0] = Session["DateFileName[0]"].ToString();
            DateFileName[1] = Session["DateFileName[1]"].ToString();*/



            try
            {
                if (SaveLocation == "")
                {
                    Response.Write("<Script language='JavaScript'>alert('請輸入存放位置！');</Script>");
                    return;
                }
                if (FTPAddress == "" || UserName == "" || PassWord == "")
                {
                    Response.Write("<Script language='JavaScript'>alert('請輸入完整資料！');</Script>");
                    return;
                }

                ftp ftpclient = new ftp(FTPAddress, UserName, PassWord);
                /*string Location = Terminal + "/" + DateFileName_Thread;
                Console.WriteLine(Location);*/
                int isSuccess = DownloadWithProgressBar(FTPAddress, UserName, PassWord, Terminal + "/" + DateFileName_Thread, SaveLocation + "\\" + DateFileName_Thread);
                Session.Add("Error", 0);
                if (isSuccess == -1)
                {
                    Response.Write("<Script language='JavaScript'>alert('下載檔案失敗，請檢查檔案是否存在！');</Script>");
                    ftpclient = null;
                    Session.Add("Error", 1);
                }
                else if (isSuccess == -2)
                {
                    Response.Write("<Script language='JavaScript'>alert('登入失敗，請檢查輸入帳號密碼是否正確！');</Script>");
                    Session.Add("Error", 1);                  
                }            
            }
            catch //(Exception ex)
            {
                //Response.Write(ex.ToString());
            }
        }

        protected void RadioBtn_Viewdetails_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
            {
                Session.Add("Number_Day_Chang", "1");
                //Number_Day_Chang = 1;
            }
            else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天") //只顯示今日資料
            {
                Session.Add("Number_Day_Chang", "0");
                //Number_Day_Chang = 0;
            }
            else
            {
                Session.Add("Number_Day_Chang", "2");
                //Number_Day_Chang = 2;
            }
            //RadioButtonList.SelectedValue
        }

        protected void Btn_UpdateInformation_Click(object sender, EventArgs e)
        {
            GridView_Detail.Visible = false;
            LbDetail.Visible = false;

            string[] DateFileName = { "", "" };

            DateFileName[0] = Session["DateFileName[0]"].ToString();
            DateFileName[1] = Session["DateFileName[1]"].ToString();

            //AnalysisFileName[0] = Session["AnalysisFileName[0]"].ToString();
            //AnalysisFileName[1] = Session["AnalysisFileName[1]"].ToString();


            //清空所有資料
            GridView_Analysis.DataBind();

            string Error;
            try
            {
                Error = Session["Error"].ToString();
            }
            catch (Exception ex)
            {
                Session.Add("Error", 0);
                Error = Session["Error"].ToString();
                Response.Write(ex.ToString());
            }

            for (int i = 0; i < 2; i++)
            {
                if (Int32.Parse(Session["Error"].ToString()) == 0)//只要其中一天讀取錯誤則直接離開(不會跳出錯誤兩次)
                {
                    if (File.Exists(SaveLocation + "\\" + DateFileName[i]))
                    {
                        //取得機器名稱，並把資料放入陣列
                        GreateMachineData(SaveLocation + "\\" + DateFileName[i], i);
                    }
                    else
                    {
                        Session.Add("Error", 1);
                        //Error = 1;
                        string[] DayList = DateFileName[i].Split('.');
                        if (i == 0)
                        {
                            Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容，並重新下載！');</Script>");
                        }
                        else if (i == 1)
                        {
                            Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容，並重新下載！');</Script>");
                        }
                        break;
                    }

                }

                else if (Int32.Parse(Session["Error"].ToString()) == 1)//只要其中一天讀取錯誤則直接離開(不會跳出錯誤兩次)
                {
                    break;
                }
            }
        }



        //private string host = null;
        //private string user = null;
        //private string pass = null;
        private FtpWebRequest ftpRequest = null;
        private FtpWebResponse ftpResponse = null;
        private Stream ftpStream = null;
        private int bufferSize = 2048;
        //public ftp(string hostIP, string userName, string password) { host = hostIP; user = userName; pass = password; }

        public int DownloadWithProgressBar(string host, string user, string pass, string remoteFile, string localFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                Console.WriteLine(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Get the FTP Server's Response Stream */
                ftpStream = ftpResponse.GetResponseStream();
                /* Open a File Stream to Write the Downloaded File */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                /* Download the File by Writing the Buffered Data Until the Transfer is Complete */
                int dataLength = (int)ftpRequest.GetResponse().ContentLength;
                /* Get the File Length */
                if (dataLength == -1) dataLength = bytesRead;
                //Set up progress bar
                /*ProcessBar.Value = 0;
                ProcessBar.Maximum = dataLength;
                lbProgress.Text = "等待下載";*/

                try
                {
                    while (bytesRead > 0)
                    {
                        localFileStream.Write(byteBuffer, 0, bytesRead);
                        bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                        /*Application.DoEvents();
                        if (ProcessBar.Value + bytesRead <= ProcessBar.Maximum)
                        {
                            ProcessBar.Value += bytesRead;
                           // lbProgress.Text = ProcessBar.Value.ToString() + "/" + dataLength.ToString();

                            ProcessBar.Refresh();
                            Application.DoEvents();
                        }*/
                    }
                    if (bytesRead == 0)
                    {
                        /*ProcessBar.Value = ProcessBar.Maximum;
                        //lbProgress.Text = dataLength.ToString() + "/" + dataLength.ToString();

                        Application.DoEvents();*/
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /*lbProgress.Text = "下載成功";
                ProcessBar.Value = 0;*/
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                if (ex.ToString().IndexOf("(530) 未登入") != -1) return -2; //登入失敗
                else return -1;  //找不到檔案
            }

        }





        class ftp
    {
        private string host = null;
        private string user = null;
        private string pass = null;
        private FtpWebRequest ftpRequest = null;
        private FtpWebResponse ftpResponse = null;
        private Stream ftpStream = null;
        private int bufferSize = 2048;

        /* Construct Object */
        public ftp(string hostIP, string userName, string password) { host = hostIP; user = userName; pass = password; }

        /* Download File */
        public void download(string remoteFile, string localFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Get the FTP Server's Response Stream */
                ftpStream = ftpResponse.GetResponseStream();
                /* Open a File Stream to Write the Downloaded File */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                /* Download the File by Writing the Buffered Data Until the Transfer is Complete */
                try
                {
                    while (bytesRead > 0)
                    {
                        localFileStream.Write(byteBuffer, 0, bytesRead);
                        bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return;
        }

        /* Download File With progressBar*/
        //public int DownloadWithProgressBar(string remoteFile, string localFile, ProgressBar ProcessBar, Label lbProgress)
        public int DownloadWithProgressBar(string remoteFile, string localFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                Console.WriteLine(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Get the FTP Server's Response Stream */
                ftpStream = ftpResponse.GetResponseStream();
                /* Open a File Stream to Write the Downloaded File */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                /* Download the File by Writing the Buffered Data Until the Transfer is Complete */
                int dataLength = (int)ftpRequest.GetResponse().ContentLength;
                /* Get the File Length */
                if (dataLength == -1) dataLength = bytesRead;
                //Set up progress bar
                /*ProcessBar.Value = 0;
                ProcessBar.Maximum = dataLength;
                lbProgress.Text = "等待下載";*/

                try
                {
                    while (bytesRead > 0)
                    {
                        localFileStream.Write(byteBuffer, 0, bytesRead);
                        bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                        /*Application.DoEvents();
                        if (ProcessBar.Value + bytesRead <= ProcessBar.Maximum)
                        {
                            ProcessBar.Value += bytesRead;
                           // lbProgress.Text = ProcessBar.Value.ToString() + "/" + dataLength.ToString();

                            ProcessBar.Refresh();
                            Application.DoEvents();
                        }*/
                    }
                    if (bytesRead == 0)
                    {
                        /*ProcessBar.Value = ProcessBar.Maximum;
                        //lbProgress.Text = dataLength.ToString() + "/" + dataLength.ToString();

                        Application.DoEvents();*/
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /*lbProgress.Text = "下載成功";
                ProcessBar.Value = 0;*/
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                return 1;
            }
            catch (Exception ex) { 
                Console.WriteLine(ex.ToString());
                if (ex.ToString().IndexOf("(530) 未登入") != -1) return -2; //登入失敗
                else return -1;  //找不到檔案
            }

        }

        /* Upload File */
        public void upload(string remoteFile, string localFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpRequest.GetRequestStream();
                /* Open a File Stream to Read the File for Upload */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesSent = localFileStream.Read(byteBuffer, 0, bufferSize);
                /* Upload the File by Sending the Buffered Data Until the Transfer is Complete */
                try
                {
                    while (bytesSent != 0)
                    {
                        ftpStream.Write(byteBuffer, 0, bytesSent);
                        bytesSent = localFileStream.Read(byteBuffer, 0, bufferSize);
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpRequest = null;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return;
        }

        /* Delete File */
        public void delete(string deleteFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + deleteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return;
        }

        /* Rename File */
        public void rename(string currentFileNameAndPath, string newFileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + currentFileNameAndPath);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.Rename;
                /* Rename the File */
                ftpRequest.RenameTo = newFileName;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return;
        }

        /* Create a New Directory on the FTP Server */
        public void createDirectory(string newDirectory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + newDirectory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return;
        }

        /* Get the Date/Time a File was Created */
        public string getFileCreatedDateTime(string fileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + fileName);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.GetDateTimestamp;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string fileInfo = null;
                /* Read the Full Response Stream */
                try { fileInfo = ftpReader.ReadToEnd(); }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return File Created Date Time */
                return fileInfo;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return "";
        }

        /* Get the Size of a File */
        public string getFileSize(string fileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + fileName);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.GetFileSize;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string fileInfo = null;
                /* Read the Full Response Stream */
                try { while (ftpReader.Peek() != -1) { fileInfo = ftpReader.ReadToEnd(); } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return File Size */
                return fileInfo;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return "";
        }

        /* List Directory Contents File/Folder Name Only */
        public string[] directoryListSimple(string directory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        /* List Directory Contents in Detail (Name, Size, Created, etc.) */
        public string[] directoryListDetailed(string directory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        public void AllSample()
        {
            /* Create Object Instance */
            ftp ftpClient = new ftp(@"ftp://10.10.10.10/", "user", "password");

            /* Upload a File */
            ftpClient.upload("etc/test.txt", @"C:\Users\metastruct\Desktop\test.txt");

            /* Download a File */
            ftpClient.download("etc/test.txt", @"C:\Users\metastruct\Desktop\test.txt");

            /* Delete a File */
            ftpClient.delete("etc/test.txt");

            /* Rename a File */
            ftpClient.rename("etc/test.txt", "test2.txt");

            /* Create a New Directory */
            ftpClient.createDirectory("etc/test");

            /* Get the Date/Time a File was Created */
            string fileDateTime = ftpClient.getFileCreatedDateTime("etc/test.txt");
            Console.WriteLine(fileDateTime);

            /* Get the Size of a File */
            string fileSize = ftpClient.getFileSize("etc/test.txt");
            Console.WriteLine(fileSize);

            /* Get Contents of a Directory (Names Only) */
            string[] simpleDirectoryListing = ftpClient.directoryListDetailed("/etc");
            for (int i = 0; i < simpleDirectoryListing.Count(); i++) { Console.WriteLine(simpleDirectoryListing[i]); }

            /* Get Contents of a Directory with Detailed File/Directory Info */
            string[] detailDirectoryListing = ftpClient.directoryListDetailed("/etc");
            for (int i = 0; i < detailDirectoryListing.Count(); i++) { Console.WriteLine(detailDirectoryListing[i]); }
            /* Release Resources */
            ftpClient = null;
        }
    }
    }
}