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


        //
        public int Over_Capacity_Number = 1000;

        //存放每一個機台下的詳細資料
        public MachineInformation[] MachineDetail = null;  //(紀錄當天資料)
        public old_MachineInformation[] old_MachineDetail = null; //(紀錄前一天資料)

        string[] strDateYear = { "", "" }; //放置年(昨天 今天)
        string[] strDateMonth = { "", "" };//放置月(昨天 今天)
        string[] strDateDay = { "", "" };//放置日(昨天 今天)

        string MachineList_SAVE;  //機台名稱暫存空間(尚未拆解字串)

        //FileStream file;

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
                //1A廠
                FTPAddress = "ftp://10.1.12.100";
                UserName = "KVIE";
                PassWord = "1234";
                SaveLocation = @"C://WebData//生產資料_1A";    //@"C://生產資料_A廠";
                Terminal = "/MMC/OOATWTN1A";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,11B06,14B19,11B01,11B09,11B02,11B10,14B03,14B09,14B08,14B12,11B14,14B22,11B03,NF120,14B11,14B20,14B16,14B15,11B05,11B04,14B04,14B21,17B02,17B01,14B24,19B34,19B33,19B32,19B31,19B30,19B29";
                

                //6A12_6B1廠
                /*FTPAddress = "ftp://10.1.12.171";
                UserName = "KVIE";
                PassWord = "1234";
                SaveLocation = @"C://WebData//生產資料_6A12_6B1";
                Terminal = "/MMC/OOATWTN6A12B1";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,08B01,27B06,19B05,19B06,24B12,27B04,27B01,14B06,14B23,14B17,14B18,14B13,14B07,11B07,14B05,75D01,75D02,15B01,14B25,14B26,11B11,11B12,11B13,14B10";
                */

                //3A14_2C13廠
                /*FTPAddress = "ftp://10.1.14.250";
                UserName = "KVIE";
                PassWord = "1234";
                SaveLocation = @"C://WebData//生產資料_3A14_2C13";
                Terminal = "/MMC/OOATWTN3A142C13";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,560P1,PF680,550P1,570P1,36B01,36B03,33B01,41B01,150DL,19B07,19B08,19B09,19B11,19B13,24B11,19B26,19B25,22B01,19B24,19B22,24B13,30B01,30B02";
                */

                //CHC中洲廠
                /*FTPAddress = "ftp://10.1.104.200";
                UserName = "KVIE";
                PassWord = "1234";
                SaveLocation = @"C://WebData//生產資料_CHC中洲";
                Terminal = "/MMC/OOATWTNCHC01";
                MachineList_SAVE = "Year,month,day,hour,minute,--------,525B01,508NF03,508NF02,508NF01,11B27,11B26,11B25,11B24,11B23,11B16,11B15,11B14,14B29,14B30,14B36,14B37,11B38,11B39,15B01,14B40,14B41,14B42,14B43,14B44,14B22,14B35,14B34,14B35,11B17,14B45";
                */
            }
        }


        protected void dateTimeAnalysis_SelectionChanged(object sender, EventArgs e)
        {
            int i = 0, day_1, day_2;
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

            Label_Yesterday_Number.Visible = false;
            Label_Today_Number.Visible = false;

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

            if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
            {
                day_1 = 1;
                day_2 = 2;
            }
            else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天") //只顯示昨天資料
            {
                day_1 = 0;
                day_2 = 1;
            }
            else  //顯示兩天資料
            {
                day_1 = 0;
                day_2 = 2;
            }

            for (i = day_1; i < day_2; i++)
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
            int i = 0, day_1, day_2;
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

                if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
                {
                    day_1 = 1;
                    day_2 = 2;
                }
                else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天") //只顯示昨天資料
                {
                    day_1 = 0;
                    day_2 = 1;
                }
                else  //顯示兩天資料
                {
                    day_1 = 0;
                    day_2 = 2;
                }

                //下載兩天資料
                for (i = day_1; i < day_2; i++)
                {
                    //ftp ftpclient = new ftp(FTPAddress, UserName, PassWord);
                    string Location = Terminal + "/" + DateFileName[i];

                    //確定可以下載
                    /*WebClient request = new WebClient();
                    request.Credentials = new NetworkCredential(UserName, PassWord);
                    byte[] fildData = request.DownloadData(FTPAddress + Terminal + "/" + DateFileName[i]);
                    file = File.Create(SaveLocation + "/" + DateFileName[i]); 
                    file.Write(fildData, 0, fildData.Length); 
                    file.Close();*/



                    int isSuccess = Download(FTPAddress + Terminal + "/" + DateFileName[i], SaveLocation, UserName, PassWord);

                    //int isSuccess = DownloadWithProgressBar(FTPAddress ,UserName, PassWord , Terminal + "/" + DateFileName[i], SaveLocation + "\\" + DateFileName[i]);
                    //int isSuccess = 0;
                    if (isSuccess == -1)
                    {
                        if (i == 0)
                        {
                            Response.Write("<Script language='JavaScript'>alert('昨日檔案下載失敗，請檢查檔案是否存在！');</Script>");
                        }
                        else if (i == 1)
                        {
                            Response.Write("<Script language='JavaScript'>alert('今日檔案下載失敗，請檢查檔案是否存在！');</Script>");
                        }
                        //Response.Write("<Script language='JavaScript'>alert('下載檔案失敗，請檢查檔案是否存在！');</Script>");
                        //ftpclient = null;
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

        int Problem = 0;
        //計算每一機台的整天統計資料
        private void GreateMachineData(string AnalysisFileName, int input_day)  //input_day = 0為昨天資料、input_day = 1為今天資料
        {
            //今日錯誤行數
            int Problem_Today = 0;

            //昨日錯誤行數
            int Problem_Yesterday = 0;

            //今日錯誤統計
            List<int> Problem_Num_Today = new List<int>();
            List<String> Problem_Num_Today_MachineName = new List<String>();

            //昨日錯誤統計
            List<int> Problem_Num_Yesterday = new List<int>();
            List<String> Problem_Num_Yesterday_MachineName = new List<String>();

            int DisplayErrorYesterday = 0;
            int DisplayErrorToday = 0;

            //錯誤計數
            int Problem_Count = 0;

            //顯示錯誤行數
            int ErrorNumberRows;

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
            
            //自行判斷資料數量使用
            int ReadLine = 0;

            DataHeight = 2000;

            try
            {
                //讀取檔案
                System.IO.StreamReader file = new System.IO.StreamReader(AnalysisFileName);

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
                //
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
                    Label_Yesterday_Number.Visible = true;
                    ALLData_old = new string[DataHeight][];
                    Label_Yesterday_Number.Text = "昨天行數：" + DataHeight.ToString(); 
                }
                else if (input_day == 1)
                {
                    Label_Today_Number.Visible = true;
                    ALLData = new string[DataHeight][];
                    Label_Today_Number.Text = "今天行數：" + DataHeight.ToString(); 
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
                Problem = 0;
                //把資料放入陣列裡
                for (int i = 0; i < DataHeight; i++)
                {
                    try
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
                        else
                        {
                            EachLine = file.ReadLine();
                            if (EachLine == null)//當讀取不到值時則跳出
                            {
                                break;
                            }
                            ReadLine++;//只要有讀取到資料則累計計數
                            //二維陣列的宣告
                            if (input_day == 0)//(前一天資料)
                            {
                                ALLData_old[i - DeleteValue_Count] = new string[DataWidth];
                            }
                            else if (input_day == 1)//(當天資料)
                            {
                                ALLData[i - DeleteValue_Count] = new string[DataWidth];
                            }
                            //string eee;
                            for (int j = 0; j < DataWidth; j++)
                            {
                                string[] SplitEachLine = EachLine.Split(',');

                                if (SplitEachLine[5] != "--------")
                                {
                                    j = DataWidth; //直接丟一個超過陣列大小的值，讓程式錯誤跳至catch內
                                }

                                if (input_day == 0)//(前一天資料)
                                {
                                    ALLData_old[i - DeleteValue_Count - Problem][j] = SplitEachLine[j]; //如果有跳錯誤 錯誤行數 直接減掉
                                    Session.Add("ALLData_old[" + (i - DeleteValue_Count).ToString() + "][" + j.ToString() + "]".ToString(), SplitEachLine[j]);
                                }
                                else if (input_day == 1)//(當天資料)
                                {
                                    ALLData[i - DeleteValue_Count - Problem][j] = SplitEachLine[j];//如果有跳錯誤 錯誤行數 直接減掉
                                    Session.Add("ALLData[" + (i - DeleteValue_Count).ToString() + "][" + j.ToString() + "]".ToString(), SplitEachLine[j]);
                                }
                            }
                            ProcessCounter++;
                        }
                    }
                    catch (Exception)
                    {
                        if (input_day == 0)//(前一天資料)
                        {
                            string Problem_Srting;
                            Problem_Srting = "<Script language='JavaScript'>alert('昨天檔案第" + (i + 1) + "行出現問題，已跳過不處理，請確認昨日檔案內容！');</Script>";
                            //Response.Write("<Script language='JavaScript'>alert('昨天檔案第%d行出現問題，請確認昨日檔案內容！');</Script>");
                            Response.Write(Problem_Srting);
                        }
                        else if (input_day == 1)//(當天資料)
                        {
                            string Problem_Srting;
                            Problem_Srting = "<Script language='JavaScript'>alert('今天檔案第" + (i + 1) + "行出現問題，已跳過不處理，請確認今日檔案內容！');</Script>";
                            //Response.Write("<Script language='JavaScript'>alert('今天檔案第%d行出現問題，請確認今日檔案內容！');</Script>");
                            Response.Write(Problem_Srting);
                        }
                        Problem++;
                        //DeleteValue_Count++;
                        ProcessCounter++;
                    }
                }

                DataHeight = ReadLine - Problem;//跳出後將計數資料放至須跑之資料內
                if (input_day == 0)//判斷為昨天資料
                {
                    //DataHeight_yesterday = ReadLine;//跳出後將計數資料放至昨天資料內
                    Session.Add("DataHeight_yesterday", DataHeight);//計算昨天資料數量(可防止昨天與今天資料量不同也可進行分析)  
                }
                else if (input_day == 1)//判斷為今天資料
                {
                    //DataHeight_today = ReadLine;//跳出後將計數資料放至今天資料內
                    Session.Add("DataHeight_today", DataHeight);//計算今天資料數量(可防止昨天與今天資料量不同也可進行分析)
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
                        for (int j = 0; j < DataHeight - 1; j++)
                        {
                            Problem_Yesterday = j;
                            try
                            {                             
                                //取得當前產量與下一分鐘的產量
                                int ThisValue = Convert.ToInt32(ALLData_old[j][old_MachineDetail[i].Index]);
                                int NextValue = Convert.ToInt32(ALLData_old[j + 1][old_MachineDetail[i].Index]);
                                //每分鐘產量>10顆
                                if (NextValue - ThisValue > 10 && Math.Abs(NextValue - ThisValue) < Over_Capacity_Number)
                                {

                                    int ThisTime_old = (Convert.ToInt32(ALLData_old[j][3]) * 60) + Convert.ToInt32(ALLData_old[j][4]);
                                    int NextTime_old = (Convert.ToInt32(ALLData_old[j + 1][3]) * 60) + Convert.ToInt32(ALLData_old[j + 1][4]);

                                    //如果下一分鐘時間比現在時間大(不可能 資料有問題)
                                    if (ThisTime_old > NextTime_old)
                                    {
                                        Problem_Num_Yesterday.Add(Problem_Yesterday);
                                        Problem_Num_Yesterday_MachineName.Add(MachineList[i + 6]);
                                        Problem_Count++;
                                        j++;
                                    }
                                    //總產能與總運作時間
                                    else
                                    {
                                        old_MachineDetail[i].TotalProduct += (NextValue - ThisValue);
                                        old_MachineDetail[i].TotalWorkTimes += (NextTime_old - ThisTime_old);
                                    }
                                }
                                else if (Math.Abs(NextValue - ThisValue) >= Over_Capacity_Number)
                                {
                                    Problem_Num_Yesterday.Add(Problem_Yesterday);
                                    Problem_Num_Yesterday_MachineName.Add(MachineList[i + 6]);
                                    Problem_Count++;
                                    j++;
                                }
                            }
                            catch (Exception)//若出現錯誤則記錄錯誤行數
                            {
                                Problem_Num_Yesterday.Add(Problem_Yesterday);
                                Problem_Num_Yesterday_MachineName.Add(MachineList[i + 6]);
                                Problem_Count++;
                                j++;
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
                        for (int j = 0; j < DataHeight - 1; j++)
                        {
                            //取得當前產量與下一分鐘的產量
                            Problem_Today = j;
                            try
                            {
                                int ThisValue = Convert.ToInt32(ALLData[j][MachineDetail[i].Index]);
                                int NextValue = Convert.ToInt32(ALLData[j + 1][MachineDetail[i].Index]);
                                //每分鐘產量>10顆
                                if (NextValue - ThisValue > 10 && Math.Abs(NextValue - ThisValue) < Over_Capacity_Number)
                                {
                                    int ThisTime = (Convert.ToInt32(ALLData[j][3]) * 60) + Convert.ToInt32(ALLData[j][4]);
                                    int NextTime = (Convert.ToInt32(ALLData[j + 1][3]) * 60) + Convert.ToInt32(ALLData[j + 1][4]);

                                    //如果下一分鐘時間比現在時間大(不可能 資料有問題)
                                    if (ThisTime > NextTime)
                                    {
                                        Problem_Num_Today.Add(Problem_Today);
                                        Problem_Num_Today_MachineName.Add(MachineList[i + 6]);
                                        Problem_Count++;
                                        j++;
                                    }
                                    else
                                    {
                                        //總產能與總運作時間
                                        MachineDetail[i].TotalProduct += (NextValue - ThisValue);
                                        //MachineDetail[i].TotalWorkTimes++;
                                        MachineDetail[i].TotalWorkTimes += (NextTime - ThisTime);
                                    }
                                }
                                else if (Math.Abs(NextValue - ThisValue) >= Over_Capacity_Number)
                                {
                                    Problem_Num_Today.Add(Problem_Today);
                                    Problem_Num_Today_MachineName.Add(MachineList[i + 6]);
                                    Problem_Count++;
                                    j++;
                                }
                            }
                            catch (Exception)//若出現錯誤則記錄錯誤行數
                            {
                                Problem_Num_Today.Add(Problem_Today);
                                Problem_Num_Today_MachineName.Add(MachineList[i + 6]);
                                Problem_Count++;
                                j++;
                            }
                        }
                        //平均產能
                        if (MachineDetail[i].TotalWorkTimes != 0) MachineDetail[i].TotalAvearage = ((float)(MachineDetail[i].TotalProduct) / (float)(MachineDetail[i].TotalWorkTimes));
                        else MachineDetail[i].TotalAvearage = 0;
                    }

                    //新增至DataGridView
                    /*if (input_day == 1)//(當天資料 判斷時須拿前一天資料做整合)
                    {*/
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
                    else if (RadioBtn_Viewdetails.SelectedValue == "顯示兩天") //顯示兩天(今、昨天)資料
                    {
                        if (input_day == 1)
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
                
                //顯示錯誤行數(昨日跟今日)
                if (input_day == 0)//(昨天資料分析)
                {
                    DisplayErrorYesterday = 0;
                    int YesterdayNameData_Num = 0;//計算在錯誤之前有幾個Year的行數
                    Problem_Num_Yesterday.Sort();//將昨天有錯誤資料做排序
                    List<int> resultList = new List<int>();//開新的LIST
                    //Problem_Num_Yesterday陣列內將重複名稱的資料清除
                    foreach (int Data in Problem_Num_Yesterday)
                    {
                        if (!resultList.Contains(Data))
                        {
                            resultList.Add(Data);
                        }
                    }
                    //除了0以外的資料錯誤行數顯示出來
                    foreach (int Data in resultList)
                    {
                        if (DisplayErrorYesterday <= 10)
                        {
                            if (Data == 0)
                            {
                                ErrorNumberRows = 1;
                            }
                            else
                            {
                                ErrorNumberRows = 2;
                            }
                            string Problem_Srting;

                            //判斷在錯誤之前有幾個Year的行數
                            for (int i = 0; i < DeleteValue_Num; i++)
                            {
                                if (NameData[i] > Data)
                                {
                                    YesterdayNameData_Num++;
                                }
                            }

                            if (DisplayErrorYesterday < 10)
                            {
                                Problem_Srting = "<Script language='JavaScript'>alert('昨天檔案：機台名稱" + Problem_Num_Yesterday_MachineName[DisplayErrorYesterday] + "，第" + (Data + DeleteValue_Count + ErrorNumberRows - YesterdayNameData_Num) + "行出現資料問題，請確認昨日檔案內容！');</Script>";
                                Response.Write(Problem_Srting);
                            }
                            else if (DisplayErrorYesterday == 10)
                            {
                                Problem_Srting = "<Script language='JavaScript'>alert('錯誤過多，不再顯示錯誤！');</Script>";
                                Response.Write(Problem_Srting);
                            }

                            DisplayErrorYesterday++;
                        }
                        
                    }
                }
                else if (input_day == 1)//(今天資料分析)
                {
                    DisplayErrorToday = 0;
                    int TodayNameData_Num = 0;//計算在錯誤之前有幾個Year的行數
                    Problem_Num_Today.Sort();//將今天有錯誤資料做排序
                    List<int> resultList = new List<int>();//開新的LIST
                    //Problem_Num_Today陣列內將重複名稱的資料清除
                    foreach (int Data in Problem_Num_Today)
                    {
                        if (!resultList.Contains(Data))
                        {
                            resultList.Add(Data);
                        }
                    }
                    if (DisplayErrorToday <= 10)
                    {
                        //除了0以外的資料錯誤行數顯示出來
                        foreach (int Data in resultList)
                        {
                            if (Data == 0)
                            {
                                ErrorNumberRows = 1;
                            }
                            else
                            {
                                ErrorNumberRows = 2;
                            }
                            string Problem_Srting;

                            //判斷在錯誤之前有幾個Year的行數
                            for (int i = 0; i < DeleteValue_Num; i++)
                            {
                                if (NameData[i] > Data)
                                {
                                    TodayNameData_Num++;
                                }
                            }
                            if (DisplayErrorToday < 10)
                            {
                                Problem_Srting = "<Script language='JavaScript'>alert('今天檔案：機台名稱" + Problem_Num_Today_MachineName[DisplayErrorToday] + "，第" + (Data + DeleteValue_Count + ErrorNumberRows - TodayNameData_Num) + "行出現資料問題，請確認今日檔案內容！');</Script>";
                                Response.Write(Problem_Srting);
                            }
                            else if (DisplayErrorToday == 10)
                            {
                                Problem_Srting = "<Script language='JavaScript'>alert('錯誤過多，不再顯示錯誤！');</Script>";
                                Response.Write(Problem_Srting);
                            }
                            DisplayErrorToday++;
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
                        string Problem_Srting;
                        Problem_Srting = "<Script language='JavaScript'>alert('昨天檔案第" + (Problem_Yesterday + DeleteValue_Count + 2) + "行出現非常嚴重資料錯誤，請確認昨日檔案內容！');</Script>";
                        Response.Write(Problem_Srting);
                        //Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容！');</Script>");
                    }
                    if (input_day == 1)
                    {
                        string Problem_Srting;
                        Problem_Srting = "<Script language='JavaScript'>alert('今天檔案第" + (Problem_Today + DeleteValue_Count + 2) + "行出現非常嚴重資料錯誤，請確認今日檔案內容！');</Script>";
                        Response.Write(Problem_Srting);
                        //Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容！');</Script>");
                    }
                }
                else
                {
                    if (input_day == 0)
                    {
                        string Problem_Srting;
                        Problem_Srting = "<Script language='JavaScript'>alert('昨天檔案出現非常嚴重資料錯誤，請確認昨日檔案內容，是否有被其他程式開啟！');</Script>";
                        Response.Write(Problem_Srting);
                        //Response.Write("<Script language='JavaScript'>alert('昨天檔案有問題，請確認昨日檔案內容！');</Script>");
                    }
                    if (input_day == 1)
                    {
                        string Problem_Srting;
                        Problem_Srting = "<Script language='JavaScript'>alert('今天檔案出現非常嚴重資料錯誤，請確認今日檔案內容，是否有被其他程式開啟！');</Script>";
                        Response.Write(Problem_Srting);
                        //Response.Write("<Script language='JavaScript'>alert('今天檔案有問題，請確認今日檔案內容！');</Script>");
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

                    DataHeight = Convert.ToInt32(Session["DataHeight_yesterday"].ToString()) ;   //將昨天之計數回傳至DataHeight內(可防止昨天與今天資料量不同也可進行分析)
                    //計算產量
                    for (int i = 1; i < DataHeight - 1; i++)
                    {
                        int ThisValue = Convert.ToInt32(Session["ALLData_old[" + i + "][" + ComboTag + "]"].ToString());
                        int NextValue = Convert.ToInt32(Session["ALLData_old[" + (i + 1) + "][" + ComboTag + "]"].ToString());

                        //每分鐘產量>100顆
                        if (NextValue - ThisValue > ProductThreshold && Math.Abs(NextValue - ThisValue) < Over_Capacity_Number)
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
                        else if (Math.Abs(NextValue - ThisValue) >= Over_Capacity_Number)//前後產能大於設定值 則不記入記數
                        {
 
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

                    DataHeight = Convert.ToInt32(Session["DataHeight_today"].ToString()) ;   //將今天之計數回傳至DataHeight內(可防止昨天與今天資料量不同也可進行分析)

                    //計算產量
                    for (int i = 1; i < DataHeight - 1; i++)
                    {
                        //取得當前產量與下一分鐘的產量
                        int ThisValue = Convert.ToInt32(Session["ALLData[" + i + "][" + ComboTag + "]"].ToString());
                        int NextValue = Convert.ToInt32(Session["ALLData[" + (i + 1) + "][" + ComboTag + "]"].ToString());
                        //每分鐘產量>100顆
                        if (NextValue - ThisValue > ProductThreshold && Math.Abs(NextValue - ThisValue) < Over_Capacity_Number)
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
                        else if (Math.Abs(NextValue - ThisValue) >= Over_Capacity_Number) //前後產能大於設定值 則不記入記數
                        {

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
                    Download_ftp();
                }
                Thread.Sleep(1000);
            }
        }

        private void Download_ftp()
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

                //ftp ftpclient = new ftp(FTPAddress, UserName, PassWord);
                /*string Location = Terminal + "/" + DateFileName_Thread;
                Console.WriteLine(Location);*/
                int isSuccess = Download(FTPAddress + Terminal + "/" + DateFileName_Thread, SaveLocation, UserName, PassWord);
                //int isSuccess = 0; DownloadWithProgressBar(FTPAddress, UserName, PassWord, Terminal + "/" + DateFileName_Thread, SaveLocation + "\\" + DateFileName_Thread);
                Session.Add("Error", 0);
                if (isSuccess == -1)
                {
                    Response.Write("<Script language='JavaScript'>alert('下載檔案失敗，請檢查檔案是否存在！');</Script>");
                    //ftpclient = null;
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

            int i = 0, day_1, day_2;
            GridView_Detail.Visible = false;
            LbDetail.Visible = false;

            string[] DateFileName = { "", "" };

            DateFileName[0] = Session["DateFileName[0]"].ToString();
            DateFileName[1] = Session["DateFileName[1]"].ToString();

            //AnalysisFileName[0] = Session["AnalysisFileName[0]"].ToString();
            //AnalysisFileName[1] = Session["AnalysisFileName[1]"].ToString();

            //清空所有資料
            GridView_Analysis.DataBind();

            Label_Yesterday_Number.Visible = false;
            Label_Today_Number.Visible = false;

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

            if (RadioBtn_Viewdetails.SelectedValue == "顯示今天") //只顯示今日資料
            {
                day_1 = 1;
                day_2 = 2;
            }
            else if (RadioBtn_Viewdetails.SelectedValue == "顯示昨天") //只顯示昨天資料
            {
                day_1 = 0;
                day_2 = 1;
            }
            else  //顯示兩天資料
            {
                day_1 = 0;
                day_2 = 2;
            }

            for (i = day_1; i < day_2; i++)
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

        /// 
        /// 從ftp下載檔案
        /// 
        /// 下載完整網址路徑ex : ftp//127.0.0.1/abc.xml 
        /// 本機存檔目錄路徑
        /// 使用者FTP登入帳號
        /// 使用者登入密碼
        /// 
        Stream responseStream = null;
        FileStream fileStream = null;
        StreamReader reader = null;
        internal int Download(string downloadUrl, string TargetPath, string UserName, string Password)
        {        
            try
            {
                FtpWebRequest downloadRequest = (FtpWebRequest)WebRequest.Create(downloadUrl);
                downloadRequest.Method = WebRequestMethods.Ftp.DownloadFile; //設定Method下載檔案
                if (UserName.Length > 0)//如果需要帳號登入
                {
                    NetworkCredential nc = new NetworkCredential(UserName, Password);
                    downloadRequest.Credentials = nc;//設定帳號
                }
                FtpWebResponse downloadResponse = (FtpWebResponse)downloadRequest.GetResponse();
                responseStream = downloadResponse.GetResponseStream();//取得FTP伺服器回傳的資料流
                string fileName = Path.GetFileName(downloadRequest.RequestUri.AbsolutePath);
                if (fileName.Length == 0)
                {
                    reader = new StreamReader(responseStream);
                    throw new Exception(reader.ReadToEnd());
                }
                else
                {
                    fileStream = File.Create(TargetPath + @"\" + fileName);
                    byte[] buffer = new byte[1024];
                    int bytesRead;
                    while (true)
                    {//開始將資料流寫入到本機
                        bytesRead = responseStream.Read(buffer, 0, buffer.Length);
                        if (bytesRead == 0)
                            break;
                        fileStream.Write(buffer, 0, bytesRead);
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                if (ex.ToString().IndexOf("(530) 未登入") != -1) return -2; //登入失敗
                else return -1;  //找不到檔案
                //throw new Exception(ex.Message);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                else if (responseStream != null)
                    responseStream.Close();
                if (fileStream != null)
                    fileStream.Close();
            }
        }
    }
}