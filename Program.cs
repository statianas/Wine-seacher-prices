using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using IConfiguration = AngleSharp.IConfiguration;
using AngleSharp.Common;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Threading;
using System.Net;
using System.Text;
using System.Diagnostics;
using IronPython.Hosting;
using System.Net.Http;
using AngleSharp.Io.Network;
using CsvHelper;
using System.Collections;
using CsvHelper.Configuration;
using IronXL;
using System.Data;

IronXL.License.LicenseKey = "IRONXL.TANYA252002.24702-EEFDEB54B7-FATUKQ-KFNL25VVTEDS-52XUXW44D23K-RNTSLUS6GEAJ-JJCF6WS4JALC-CPLOOAINBCYQ-FYBHWN-TL3RNYS4T6SHEA-DEPLOYMENT.TRIAL-3KK4ZX.TRIAL.EXPIRES.17.SEP.2022";
DataTable ReadExcel(string fileName)
{
    WorkBook workbook = WorkBook.Load(fileName);
    WorkSheet sheet = workbook.DefaultWorkSheet;
    return sheet.ToDataTable(true);
}

DataTable ReadCSVData(string csvFileName)
{
    var csvFilereader = new DataTable();
    csvFilereader = ReadExcel(csvFileName);
    return csvFilereader;
}

var csvFilereader = ReadCSVData("data2.csv");



//using AngleSharp.Io.Network;
//using AngleSharp.Network;

string pathPageLogin = "https://www.wine-searcher.com/sign-in?pro_redirect_url_F=%2F";



//string url = "https://www.wine-searcher.com/find/marchesi+antinori+badia+a+passignano+grand+select+docg+chianti+cls+tuscany+italy/1/italy";
//string url = "";




/* для получени прокси с кода на питоне

//ProcessStartInfo startInfo = new ProcessStartInfo("C:/Users/Таня/AppData/Local/Programs/Python/Python39/python.exe");
ProcessStartInfo startInfo = new ProcessStartInfo("C:/Users/Таня/PycharmProjects/parser_example/venv/Scripts/python.exe");

Process process = new Process();

string directory = @"C:\\Users\\Таня\\PycharmProjects\\parser_example";
string script = "for_proxy.py";

startInfo.WorkingDirectory = directory;
startInfo.Arguments = script;
startInfo.UseShellExecute = false;
startInfo.CreateNoWindow = true;
startInfo.RedirectStandardError = true;
startInfo.RedirectStandardOutput = true;

process.StartInfo = startInfo;

process.Start();
await Task.Delay(60000);
string output = process.StandardOutput.ReadToEnd();
process.Close();

Console.Write(output);
Console.Write("lol");
*/

/*
 
var handler = new HttpClientHandler
{
    Proxy = new WebProxy(String.Format("{0}:{1}", "159.197.250.171", "3128"), false),//false
    PreAuthenticate = true,//true
    UseDefaultCredentials = false,//false
    UseProxy = true,
    UseCookies = false,
    AllowAutoRedirect = false,

};

*/

var config = Configuration.Default.WithDefaultLoader().WithCookies();
 
//с прокси

/*

var handler = new HttpClientHandler
{
    Proxy = new WebProxy(String.Format("{0}:{1}", "159.197.250.142", "3128"), false),//false

    UseDefaultCredentials = false,//false
    UseProxy = true,
    UseCookies = false,//false
    AllowAutoRedirect = false,//nin
    PreAuthenticate = true,


};


var client = new HttpClient(handler);
client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36");
client.DefaultRequestHeaders.Add("Accept-Language", "en-US");//nin
var requester = new HttpClientRequester(client);

var config = Configuration.Default
  .With(requester)
  .WithJs()
  .WithDefaultLoader()
  .WithTemporaryCookies();

//WithRequesters()
*/

Thread.Sleep(3000);
IBrowsingContext browsingContext = BrowsingContext.New(config);
IDocument queryDocument = await browsingContext.OpenAsync(pathPageLogin);
Thread.Sleep(3000);


//Console.Write(browsingContext.Active.Title);

//Console.Write(queryDocument.DocumentElement.OuterHtml);


//авторизация


browsingContext.Active.QuerySelector<IHtmlInputElement>("#loginmodel-username").Value = "antonskrobotov@gmail.com";
Thread.Sleep(1000);
browsingContext.Active.QuerySelector<IHtmlInputElement>("#loginmodel-password").Value = "s64_ftYn8";
Thread.Sleep(1000);
var form = queryDocument.QuerySelector<IHtmlFormElement>("#loginsmallform");
var resultDocument = await form.SubmitAsync();



//await browsingContext.OpenAsync(url1);
//await browsingContext.OpenAsync(url);
//resultDocument = await browsingContext.OpenAsync(url);


string fourth = "";
int N = 262;

for (int i = 171; i< N; i++)
{
    WorkBook wb = WorkBook.Load("file.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows[i+1].Columns[0].Value = csvFilereader.Rows[i][0].ToString();
    ws.Rows[i+1].Columns[1].Value = csvFilereader.Rows[i][1].ToString();
    string link_ws = csvFilereader.Rows[i][2].ToString();
    ws.Rows[i+1].Columns[2].Value = link_ws;
    await browsingContext.OpenAsync(link_ws); 

    //https://spectrox.ru/strikethrough/
    try
    {
        fourth = browsingContext.Active.QuerySelectorAll("script")[3].Text();//3
    }
    catch (System.ArgumentOutOfRangeException e) {

        ws.Rows[i+1].Columns[3].Value = "---";
        wb.SaveAs("file.xlsx");
        Thread.Sleep(360000);
        config = Configuration.Default.WithDefaultLoader().WithCookies();
        browsingContext = BrowsingContext.New(config);
        queryDocument = await browsingContext.OpenAsync(pathPageLogin);
        Thread.Sleep(3000);
        browsingContext.Active.QuerySelector<IHtmlInputElement>("#loginmodel-username").Value = "antonskrobotov@gmail.com";
        Thread.Sleep(1000);
        browsingContext.Active.QuerySelector<IHtmlInputElement>("#loginmodel-password").Value = "s64_ftYn8";
        Thread.Sleep(1000);
        form = queryDocument.QuerySelector<IHtmlFormElement>("#loginsmallform");
        resultDocument = await form.SubmitAsync();
        continue;
    }
    if (fourth.Length == 0) fourth = "---";
    ws.Rows[i+1].Columns[3].Value = fourth;
    wb.SaveAs("file.xlsx");
    Console.Write("Step num ", i);
    Thread.Sleep(360000);
    }
return 0;

string url2 = "https://www.wine-searcher.com/find/vignoble+de+verdot+clos+blanc+moelleux+bergerac+south+west+france/2012#t4";
Thread.Sleep(1000);
await browsingContext.OpenAsync(url2);
//Console.Write(browsingContext.Active.DocumentElement.OuterHtml);
Thread.Sleep(1000);
Console.Write(browsingContext.Active.QuerySelectorAll("script")[3].Text());
Console.Write("First");
Thread.Sleep(360000);
url2 = "https://www.wine-searcher.com/find/krug+clos+d+ambonnay+blanc+de+noir+brut+champagne+france/2002#t4";
await browsingContext.OpenAsync(url2);
Thread.Sleep(1000);
Console.Write(browsingContext.Active.QuerySelectorAll("script")[3].Text());
Console.Write("Second");
Thread.Sleep(360000);
url2 = "https://www.wine-searcher.com/find/f+bergeron+marion+clos+de+blanc+noir+brut+champagne+france/2007#t4";
await browsingContext.OpenAsync(url2);
Thread.Sleep(1000);
Console.Write(browsingContext.Active.QuerySelectorAll("script")[3].Text());
Console.Write("Third");
Thread.Sleep(360000);
url2 = "https://www.wine-searcher.com/find/les+paques+cuvee+prestige+eleve+fut+de+chene+blaye+cote+bordeaux+france/2018#t4";
await browsingContext.OpenAsync(url2);
Thread.Sleep(1000);
Console.Write(browsingContext.Active.QuerySelectorAll("script")[3].Text());
Console.Write("Fourth");



/* для проверки ip
string url2 = "https://www.myip.com/";
IDocument doc = await browsingContext.OpenAsync(url2);
Console.Write(doc.DocumentElement.OuterHtml);
*/





