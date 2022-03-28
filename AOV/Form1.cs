using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections;



namespace AOV
{

    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            



        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //編號輸入到textbox1
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("請輸入編號");
                
            }
            else
            {
                string str = textBox1.Text.ToString();
                char[] delimiterChars = { ',' };
                string[] number = textBox1.Text.Split(delimiterChars);
                /*if(textBox9 != null && textBox10 != null)
                {
                    Array.Resize(ref number, number.Length + 2);
                    number[number.Length - 2] = "998";
                    number[number.Length - 1] = "999";
                }*/
                if (number.Length < 10)
                {
                    MessageBox.Show("人數不足");
                }
                else
                {


                    int[] intnumber = number.Select(int.Parse).ToArray();


                    //亂數決定候補名單
                    Random rnd = new Random();
                    int rndx = number.Length % 5;
                    for (int random = 0; random < rndx; random++)
                    {
                        int num = rnd.Next(number.Length);
                        TextBox[] textBoxes = new TextBox[] { textBox2, textBox3, textBox4, textBox5 };
                        textBoxes[random].Text = number[num];
                        number = number.Where(val => val != number[num]).ToArray();
                    }
                    //總共隊伍數量
                    int total = number.Length / 5;
                    textBox7.Text = total.ToString();                    
                    //連結GOOGLE表單
                    UserCredential credential;
                    string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        // The file token.json stores the user's access and refresh tokens, and is created
                        // automatically when the authorization flow completes for the first time.
                        string credPath = "token.json";
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None, new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });
                    String spreadsheetId = "1m2eCxCtvnNQMGj7h3xBbTXs-hhFZ2QfTc6_L4VV3zko";//女王群
                    //String spreadsheetId = "1W7mM5yh1S6UJudk2C0_a6QTxuQCYro-ajXad78OFvSM";//怪獸社區
                    //String spreadsheetId = "1YzVQpEoc2BW83ccR-zzNTM2aA9Gkt-ArG_642LkV0uw";
                    String range = "基本資料!A2:D";
                    SpreadsheetsResource.ValuesResource.GetRequest request =
                            service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    IList<IList<Object>> values = response.Values;
                    int ranktotal = 0;
                    //計算分數用的變數
                    int rankss = 0; int ranks = 0; int rankaa = 0; int ranka = 0; int rankb = 0; int rankc = 0; int rankd = 0; int ranke = 0;
                    //記錄各排位人員的陣列
                    string[] rnkss = { }; string[] rnks = { }; string[] rnkaa = { }; string[] rnka = { }; string[] rnkb = { }; string[] rnkc = { }; string[] rnkd = { }; string[] rnke = { };
                    //分隊用變數
                    int team01 = 0; int team02 = 0; int team03 = 0; int team04 = 0; int team05 = 0; int team06 = 0; int team07 = 0; int team08 = 0; int team09 = 0; int team10 = 0;
                    int team11 = 0; int team12 = 0; int team13 = 0; int team14 = 0; int team15 = 0; int team16 = 0; int team17 = 0; int team18 = 0; int team19 = 0; int team20 = 0;
                    //分隊陣列
                    int[] allteam = { };
                    if (values != null && values.Count > 0)
                    {
                        foreach (var row in values)
                        {

                            for (int team = 0; team < number.Length; team++)
                            {
                                string str0 = row[0].ToString();//編號
                                string str1 = row[1].ToString();//遊戲名稱
                                string str2 = row[2].ToString();//line名稱
                                string str3 = row[3].ToString();//歷史最高
                                if (str0.Equals(number[team]))//編號對照
                                {
                                    //計算排位
                                    if (str3 == "SS" || str3 == "ss")
                                    {
                                        rankss += 1;
                                        Array.Resize(ref rnkss, rnkss.Length + 1);
                                        rnkss[rnkss.Length - 1] = str0;
                                    }
                                    else if (str3 == "S" || str3 == "s")
                                    {
                                        ranks += 1;
                                        Array.Resize(ref rnks, rnks.Length + 1);
                                        rnks[rnks.Length - 1] = str0;
                                    }
                                    else if (str3 == "A+1" || str3 == "A+2" || str3 == "A+3" || str3 == "A+4" || str3 == "a+1" || str3 == "a+2" || str3 == "a+3" || str3 == "a+4")
                                    {
                                        rankaa += 1;
                                        Array.Resize(ref rnkaa, rnkaa.Length + 1);
                                        rnkaa[rnkaa.Length - 1] = str0;
                                    }
                                    else if (str3 == "A+5" || str3 == "A1" || str3 == "A2" || str3 == "A3" || str3 == "A4" || str3 == "aa5" || str3 == "a1" || str3 == "a2" || str3 == "a3" || str3 == "a4")
                                    {
                                        ranka += 1;
                                        Array.Resize(ref rnka, rnka.Length + 1);
                                        rnka[rnka.Length - 1] = str0;
                                    }
                                    else if (str3 == "A5" || str3 == "B1" || str3 == "B2" || str3 == "B3" || str3 == "B4" || str3 == "a5" || str3 == "b1" || str3 == "b2" || str3 == "b3" || str3 == "b4")
                                    {
                                        rankb += 1;
                                        Array.Resize(ref rnkb, rnkb.Length + 1);
                                        rnkb[rnkb.Length - 1] = str0;
                                    }
                                    else if (str3 == "B5" || str3 == "C1" || str3 == "C2" || str3 == "C3" || str3 == "C4" || str3 == "b5" || str3 == "c1" || str3 == "c2" || str3 == "c3" || str3 == "c4")
                                    {
                                        rankc += 1;
                                        Array.Resize(ref rnkc, rnkc.Length + 1);
                                        rnkc[rnkc.Length - 1] = str0;
                                    }
                                    else if (str3 == "C5" || str3 == "D1" || str3 == "D2" || str3 == "D3" || str3 == "D4" || str3 == "c5" || str3 == "d1" || str3 == "d2" || str3 == "d3" || str3 == "d4")
                                    {
                                        rankd += 1;
                                        Array.Resize(ref rnkd, rnkd.Length + 1);
                                        rnkd[rnkd.Length - 1] = str0;
                                    }
                                    else if (str3 == "D5" || str3 == "E1" || str3 == "E2" || str3 == "E3" || str3 == "E4" || str3 == "d5" || str3 == "e1" || str3 == "e2" || str3 == "e3" || str3 == "e4")
                                    {
                                        ranke += 1;
                                        Array.Resize(ref rnke, rnke.Length + 1);
                                        rnke[rnke.Length - 1] = str0;
                                    }
                                }
                            }
                        }
                        //rank陣列亂數排序副程式
                        void Shuffle<T>(T[] Source)
                        {
                            if (Source == null) return;
                            int len = Source.Length;//用變數記會快一點點點
                            Random rd = new Random();
                            int r;//記下隨機產生的號碼
                            T tmp;//暫存用
                            for (int i = 0; i < len - 1; i++)
                            {
                                r = rd.Next(i, len);//取亂數，範圍從自己到最後，決定要和哪個位置交換，因此也不用跑最後一圈了
                                if (i == r) continue;
                                tmp = Source[i];
                                Source[i] = Source[r];
                                Source[r] = tmp;
                            }
                        }
                        Shuffle(rnkss);
                        Shuffle(rnks);
                        Shuffle(rnkaa);
                        Shuffle(rnka);
                        Shuffle(rnkb);
                        Shuffle(rnkc);
                        Shuffle(rnkd);
                        Shuffle(rnke);
                        //計算總分
                        ranktotal = rankss * 8 + ranks * 7 + +rankaa * 6 + ranka * 5 + rankb * 4 + rankc * 3 + rankd * 2 + ranke * 1;
                        textBox8.Text = ranktotal.ToString();
                        //每隊分數
                        int everyteam = ranktotal / Convert.ToInt16(textBox7.Text);
                        for(int teamup = 0;teamup < Convert.ToInt16(textBox7.Text);teamup++)
                        {
                            Array.Resize(ref allteam, allteam.Length + 1);
                            allteam[allteam.Length - 1] = Convert.ToInt16(0);
                        }
                        //各隊成員
                        string[] ateam = {  }; string[] bteam = {  }; string[] cteam = {  }; string[] dteam = {  }; string[] eteam = {  };
                        string[] fteam = {  }; string[] gteam = {  }; string[] hteam = {  }; string[] iteam = {  }; string[] jteam = {  };
                        string[] kteam = {  }; string[] lteam = {  }; string[] mteam = {  }; string[] nteam = {  }; string[] oteam = {  };
                        string[] pteam = {  }; string[] qteam = {  }; string[] rteam = {  }; string[] steam = {  }; string[] tteam = {  };
                        //4隊以內做亂數
                        if(allteam.Length <= 4 )
                        {
                            //亂數
                             teamgroupup(0);
                        }
                        else
                        {
                            //不做亂數
                             teamgroupup(1);
                        }
                        
                        //列出隊員
                        for (int print = 0; print < allteam.Length; print++)
                        {
                            switch (print)
                            {
                                case 0:
                                    textBox6.Text += "A team" +" 共"+ allteam[0] + "分" + Environment.NewLine;
                                    printteam(ateam);
                                    break;
                                case 1:
                                    textBox6.Text += Environment.NewLine + "B team" + " 共" + allteam[1] + "分" + Environment.NewLine;
                                    printteam(bteam);
                                    break;
                                case 2:
                                    textBox6.Text += Environment.NewLine + "C team" + " 共" + allteam[2] + "分" + Environment.NewLine;
                                    printteam(cteam);
                                    break;
                                case 3:
                                    textBox6.Text += Environment.NewLine + "D team" + " 共" + allteam[3] + "分" + Environment.NewLine;
                                    printteam(dteam);
                                    break;
                                case 4:
                                    textBox6.Text += Environment.NewLine + "E team" + " 共" + allteam[4] + "分" + Environment.NewLine;
                                    printteam(eteam);
                                    break;
                                case 5:
                                    textBox6.Text += Environment.NewLine + "F team" + " 共" + allteam[5] + "分" + Environment.NewLine;
                                    printteam(fteam);
                                    break;
                                case 6:
                                    textBox6.Text += Environment.NewLine + "G team" + " 共" + allteam[6] + "分" + Environment.NewLine;
                                    printteam(gteam);
                                    break;
                                case 7:
                                    textBox6.Text += Environment.NewLine + "H team" + " 共" + allteam[7] + "分" + Environment.NewLine;
                                    printteam(hteam);
                                    break;
                                case 8:
                                    textBox6.Text += Environment.NewLine + "I team" + " 共" + allteam[8] + "分" + Environment.NewLine;
                                    printteam(iteam);
                                    break;
                                case 9:
                                    textBox6.Text += Environment.NewLine + "J team" + " 共" + allteam[9] + "分" + Environment.NewLine;
                                    printteam(jteam);
                                    break;
                                case 10:
                                    textBox6.Text += Environment.NewLine + "K team" + " 共" + allteam[10] + "分" + Environment.NewLine;
                                    printteam(kteam);
                                    break;
                                case 11:
                                    textBox6.Text += Environment.NewLine + "L team" + " 共" + allteam[11] + "分" + Environment.NewLine;
                                    printteam(lteam);
                                    break;
                                case 12:
                                    textBox6.Text += Environment.NewLine + "M team" + " 共" + allteam[12] + "分" + Environment.NewLine;
                                    printteam(mteam);
                                    break;
                                case 13:
                                    textBox6.Text += Environment.NewLine + "N team" + " 共" + allteam[13] + "分" + Environment.NewLine;
                                    printteam(nteam);
                                    break;
                                case 14:
                                    textBox6.Text += Environment.NewLine + "O team" + " 共" + allteam[14] + "分" + Environment.NewLine;
                                    printteam(qteam);
                                    break;
                                case 15:
                                    textBox6.Text += Environment.NewLine + "P team" + " 共" + allteam[15] + "分" + Environment.NewLine;
                                    printteam(pteam);
                                    break;
                                case 16:
                                    textBox6.Text += Environment.NewLine + "Q team" + " 共" + allteam[16] + "分" + Environment.NewLine;
                                    printteam(qteam);
                                    break;
                                case 17:
                                    textBox6.Text += Environment.NewLine + "R team" + " 共" + allteam[17] + "分" + Environment.NewLine;
                                    printteam(rteam);
                                    break;
                                case 18:
                                    textBox6.Text += Environment.NewLine + "S team" + " 共" + allteam[18] + "分" + Environment.NewLine;
                                    printteam(steam);
                                    break;
                                case 19:
                                    textBox6.Text += Environment.NewLine + "T team" + " 共" + allteam[19] + "分" + Environment.NewLine;
                                    printteam(tteam);
                                    break;
                            }
                        }
                        //列印各隊成員副程式
                        void printteam<P>(P[] source)
                        {
                            foreach (var row in values)
                            {

                                for (int team = 0; team < 5; team++)
                                {
                                    string str0 = row[0].ToString();//編號
                                    string str1 = row[1].ToString();//遊戲名稱
                                    string str2 = row[2].ToString();//line名稱
                                    string str3 = row[3].ToString();//歷史最高

                                    if (str0.Equals(source[team]))//編號對照
                                    {

                                        //列出人員
                                        byte[] bytestr = Encoding.GetEncoding("big5").GetBytes(str1);
                                        if (bytestr.Length<8)
                                        {
                                            string teamok = str0 + "\t" + str1 + "\t" + "\t" + str2;
                                            textBox6.Text += teamok + Environment.NewLine;
                                        }
                                        else
                                        {
                                            string teamok = str0 + "\t" + str1 + "\t" + str2;
                                            textBox6.Text += teamok + Environment.NewLine;
                                        }
                                    }
                                    else { }
                                }
                            }
                        }
                        //分隊
                        void teamgroupup(int x)
                        {
                            
                            //rank ss分隊
                            for (int y = rnkss.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 8;
                                int teamnum;
                                if(x==0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnkss.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }
                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnkss[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnkss[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnkss = rnkss.Where(val => val != rnkss[teamnum]).ToArray();
                            }
                            //rank s分隊
                            for (int y = rnks.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 7;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnks.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }

                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnks[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnks[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnks = rnks.Where(val => val != rnks[teamnum]).ToArray();
                            }
                            //rank a+分隊
                            for (int y = rnkaa.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 6;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnkaa.Length);
                                }
                                else
                                {
                                    teamnum = y - 1;
                                }

                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnkaa[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnkaa = rnkaa.Where(val => val != rnkaa[teamnum]).ToArray();
                            }
                            //rank a分隊
                            for (int y = rnka.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 5;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnka.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }
                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnka[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnka[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnka = rnka.Where(val => val != rnka[teamnum]).ToArray();

                            }
                            //rank b分隊
                            for (int y = rnkb.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 4;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnkb.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }

                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnkb[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnkb[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnkb = rnkb.Where(val => val != rnkb[teamnum]).ToArray();

                            }
                            //rank c分隊
                            for (int y = rnkc.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 3;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnkc.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }
                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnkc[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnkc[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnkc = rnkc.Where(val => val != rnkc[teamnum]).ToArray();

                            }
                            //rank d分隊
                            for (int y = rnkd.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 3;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnkd.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }
                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnkd[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnkd[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnkd = rnkd.Where(val => val != rnkd[teamnum]).ToArray();

                            }
                            //rank e分隊
                            for (int y = rnke.Length; y > 0; y--)
                            {
                                int k = Array.FindIndex(allteam, val => val == allteam.Min());
                                allteam[k] += 3;
                                int teamnum;
                                if (x == 0)
                                {
                                    //亂數
                                    Random rndteam = new Random();
                                    teamnum = rndteam.Next(rnke.Length);
                                }
                                else
                                {
                                    teamnum = y-1;
                                }
                                switch (k)
                                {
                                    case 0:
                                        Array.Resize(ref ateam, ateam.Length + 1);
                                        ateam[ateam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 1:
                                        Array.Resize(ref bteam, bteam.Length + 1);
                                        bteam[bteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 2:
                                        Array.Resize(ref cteam, cteam.Length + 1);
                                        cteam[cteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 3:
                                        Array.Resize(ref dteam, dteam.Length + 1);
                                        dteam[dteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 4:
                                        Array.Resize(ref eteam, eteam.Length + 1);
                                        eteam[eteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 5:
                                        Array.Resize(ref fteam, fteam.Length + 1);
                                        fteam[fteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 6:
                                        Array.Resize(ref gteam, gteam.Length + 1);
                                        gteam[gteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 7:
                                        Array.Resize(ref hteam, hteam.Length + 1);
                                        hteam[hteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 8:
                                        Array.Resize(ref iteam, iteam.Length + 1);
                                        iteam[iteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 9:
                                        Array.Resize(ref jteam, jteam.Length + 1);
                                        jteam[jteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 10:
                                        Array.Resize(ref kteam, kteam.Length + 1);
                                        kteam[kteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 11:
                                        Array.Resize(ref lteam, lteam.Length + 1);
                                        lteam[lteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 12:
                                        Array.Resize(ref mteam, mteam.Length + 1);
                                        mteam[mteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 13:
                                        Array.Resize(ref nteam, nteam.Length + 1);
                                        nteam[nteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 14:
                                        Array.Resize(ref oteam, oteam.Length + 1);
                                        oteam[oteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 15:
                                        Array.Resize(ref pteam, pteam.Length + 1);
                                        pteam[ateam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 16:
                                        Array.Resize(ref qteam, qteam.Length + 1);
                                        qteam[qteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 17:
                                        Array.Resize(ref rteam, rteam.Length + 1);
                                        rteam[rteam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 18:
                                        Array.Resize(ref steam, steam.Length + 1);
                                        steam[steam.Length - 1] = rnke[teamnum];
                                        break;
                                    case 19:
                                        Array.Resize(ref tteam, tteam.Length + 1);
                                        tteam[tteam.Length - 1] = rnke[teamnum];
                                        break;
                                    default:
                                        break;
                                }
                                rnke = rnke.Where(val => val != rnke[teamnum]).ToArray();

                            }
                        }
                    }
                }
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            for(int i=1;i<=150;i++)
            {
                if (i != 150)
                {
                    textBox1.Text += i + ",";
                }
                else { textBox1.Text += i; }
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            

        }
    }
}
