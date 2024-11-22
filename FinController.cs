using Ruamchai.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Web;
using iText.Layout.Element;
using Paragraph = iTextSharp.text.Paragraph;
using Microsoft.Net.Http.Headers;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Ruamchai.Controllers
{
    public class FinController : Controller
    {

        // ดึงค่า connection string จาก appsettings.json
        private readonly string connectionString;
        public FinController(IConfiguration configuration)
        {
            connectionString = configuration.GetConnectionString("StrConn");
        }
        //-------------------------------------------------------------------

         
        // GET: /<controller>/
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Finreport()
        {
            return View();
        }
        
//--------------------------------------------------------------------------------- ข้อมูลสำหรับคนไข้ใน ------------------------------------------------------------------------------------------------//
        private DataTable GetexportIPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = new DataTable();

            // กำหนด format ของวันที่ที่ผู้ใช้กรอกเข้ามา
            var formatdate = $"{Printdate.Day}/{Printdate.Month}/{Printdate.Year}";
  
            string sql = "";

            // การดึงข้อมูลจาก dropdown    
            if(Selectoption == 1)      
            { 
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Doc_No ELSE NULL END AS Doc_No ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Hn ELSE NULL END AS Hn ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN YEAR(FinIpdMaster.HnYear) ELSE NULL END AS HnYear ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.An ELSE NULL END AS An ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName ,
                FinDetail.FinNameT,FinIpdDetail.Price,FinIpdDetail.DiscTotal,FinIpdDetail.Paid, FinIpdMaster.TotalNet, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN CAST(FinIpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTRIM(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTRIM(STR(MONTH(FinIpdMaster.PrintDateTime))) +'/' + LTRIM(STR(YEAR(FinIpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinIpdMaster.PrintDateTime) + ':' + DATENAME(mi,FinIpdMaster.PrintDateTime) AS Printtime , 
                FinIpdMaster.PaymentCode, FinDetail.FinCode ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay,Payment.PaymentName ORDER BY FinIpdMaster.Doc_No) = 1 THEN Payment.PaymentName ELSE '' END AS PaymentName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinIpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinIpdMaster.TotalPay) FROM FinIpdMaster WHERE FinIpdMaster.FinDocCode = ('IN1') 
                AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <> 'P' AND FinIpdMaster.PaymentCode <> 'L'  
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate)) ELSE NULL END AS TotalPrice
                FROM FinIpdMaster 
                INNER JOIN PatientData ON FinIpdMaster.Hn = PatientData.HN AND FinIpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinIpdDetail ON FinIpdMaster.Doc_No = FinIpdDetail.Doc_No AND FinIpdMaster.Doc_Yr = FinIpdDetail.Doc_Yr AND FinIpdMaster.An = FinIpdDetail.An 
                AND FinIpdMaster.AnYear = FinIpdDetail.AnYear INNER JOIN FinDetail ON FinIpdDetail.FinCode = FinDetail.FinCode
                INNER JOIN Payment ON FinIpdMaster.PaymentCode = Payment.PaymentCode
                WHERE FinIpdMaster.FinDocCode = ('IN1') 
                AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <> 'P' AND FinIpdMaster.PaymentCode <> 'L'  
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinIpdMaster.Doc_No";
            }
            else if(Selectoption == 2)               
            {
                sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Doc_No ELSE NULL END AS Doc_No ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Hn ELSE NULL END AS Hn ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN YEAR(FinIpdMaster.HnYear) ELSE NULL END AS HnYear ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.An ELSE NULL END AS An ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName ,
                FinDetail.FinNameT, FinIpdDetail.Price, FinIpdDetail.DiscTotal, FinIpdDetail.paid, FinIpdMaster.TotalNet,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No,PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN CAST(FinIpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTrim(Str(Day(FinIpdMaster.PrintDateTime))) + '/' + LTrim(Str(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinIpdMaster.PrintDateTime) + ':' +  DATENAME(mi,FinIpdMaster.PrintDateTime) AS Printtime,
                FinIpdMaster.PaymentCode, FinDetail.FinCode,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No,PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay,Companies.CompanyName ORDER BY FinIpdMaster.Doc_No) = 1 THEN Companies.CompanyName ELSE '' END AS CompanyName, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinIpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinIpdMaster.TotalPay) FROM FinIpdMaster WHERE FinIpdMaster.FinDocCode = ('IN2') AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <> 'P' AND FinIpdMaster.PaymentCode <> 'L' 
				AND LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime)))= @Printdate) ELSE NULL END AS TotalPrice
                FROM FinIpdMaster  
                INNER JOIN PatientData  ON FinIpdMaster.Hn = PatientData.Hn AND FinIpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinIpdDetail  ON FinIpdMaster.Doc_No = FinIpdDetail.Doc_No AND FinIpdMaster.Doc_Yr = FinIpdDetail.Doc_yr AND FinIpdMaster.An = FinIpdDetail.An AND FinIpdMaster.AnYear = FinIpdDetail.AnYear 
                INNER JOIN FinDetail  ON FinIpdDetail.FinCode = FinDetail.FinCode 
                INNER JOIN Companies ON FinIpdMaster.CompanyCode = Companies.CompanyCode
                WHERE FinIpdMaster.FinDocCode = ('IN2') AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <> 'P' AND FinIpdMaster.PaymentCode <> 'L' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime)))= @Printdate) 
                ORDER BY FinIpdMaster.Doc_No";            
            }                
            else if(Selectoption == 3)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Doc_No ELSE NULL END AS Doc_No ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Hn ELSE NULL END AS Hn ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN YEAR(FinIpdMaster.HnYear) ELSE NULL END AS HnYear ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.An ELSE NULL END AS An ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName ,
                FinDetail.FinNameT, FinIpdDetail.Price, FinIpdDetail.DiscTotal, FinIpdDetail.paid, FinIpdMaster.TotalNet,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN CAST(FinIpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTrim(Str(Day(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' +LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinIpdMaster.PrintDateTime) + DATENAME(mi,FinIpdMaster.PrintDateTime) AS Printtime, 
                FinIpdMaster.PaymentCode,FinDetail.Fincode ,Payment.PaymentName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinIpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinIpdMaster.TotalPay) FROM FinIpdMaster WHERE FinIpdMaster.PHeader = 'ใบแจ้งค่าใช้จ่าย' AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <>'L' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate))  ELSE NULL END AS TotalPrice
                FROM FinIpdMaster  INNER JOIN PatientData  ON FinIpdMaster.Hn = PatientData.HN 
                AND FinIpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinIpdDetail  ON FinIpdMaster.Doc_No = FinIpdDetail.Doc_No AND FinIpdMaster.Doc_Yr = FinIpdDetail.Doc_Yr AND FinIpdMaster.An = FinIpdDetail.An AND FinIpdMaster.AnYear = FinIpdDetail.AnYear 
                INNER JOIN FinDetail  ON FinIpdDetail.Fincode = FinDetail.Fincode 
                INNER JOIN Payment ON FinIpdMaster.PaymentCode = Payment.PaymentCode
                WHERE FinIpdMaster.PHeader = 'ใบแจ้งค่าใช้จ่าย' AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode <>'L' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinIpdMaster.Doc_No";               
            }               
            else if(Selectoption == 4)              
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Doc_No ELSE NULL END AS Doc_No ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Hn ELSE NULL END AS Hn ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN YEAR(FinIpdMaster.HnYear) ELSE NULL END AS HnYear ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.An ELSE NULL END AS An ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName ,
                FinDetail.FinNameT, FinIpdDetail.Price, FinIpdDetail.DiscTotal, FinIpdDetail.paid, FinIpdMaster.TotalNet,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN CAST(FinIpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTrim(Str(Day(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' +LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinIpdMaster.PrintDateTime) + DATENAME(mi,FinIpdMaster.PrintDateTime) AS Printtime,
                FinIpdMaster.PaymentCode,FinDetail.Fincode,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No,PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay,Companies.CompanyName ORDER BY FinIpdMaster.Doc_No) = 1 THEN Companies.CompanyName ELSE '' END AS CompanyName, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinIpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinIpdMaster.TotalPay) FROM FinIpdMaster WHERE FinIpdMaster.FinDocCode = ('IN3') AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode = 'L' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate) )ELSE NULL END AS TotalPrice
                FROM FinIpdMaster  
                INNER JOIN PatientData  ON FinIpdMaster.Hn = PatientData.HN AND FinIpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinIpdDetail  ON FinIpdMaster.Doc_No = FinIpdDetail.Doc_No AND FinIpdMaster.Doc_Yr = FinIpdDetail.Doc_Yr AND FinIpdMaster.An = FinIpdDetail.An AND FinIpdMaster.AnYear = FinIpdDetail.AnYear 
                INNER JOIN FinDetail  ON FinIpdDetail.Fincode = FinDetail.Fincode 
                INNER JOIN Companies ON FinIpdMaster.CompanyCode = Companies.CompanyCode
                WHERE FinIpdMaster.FinDocCode = ('IN3') AND FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.DocFlag <> 'F' AND FinIpdMaster.PaymentCode = 'L' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinIpdMaster.Doc_No";
            }              
            else if(Selectoption == 5)      
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Doc_No ELSE NULL END AS Doc_No ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.Hn ELSE NULL END AS Hn ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN YEAR(FinIpdMaster.HnYear) ELSE NULL END AS HnYear ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.Doc_No, FinIpdDetail.Doc_Yr, FinIpdMaster.Hn, YEAR(FinIpdMaster.HnYear), FinIpdMaster.An ORDER BY FinIpdMaster.Doc_No) = 1 THEN FinIpdMaster.An ELSE NULL END AS An ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName ,FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName ,
                FinDetail.FinNameT, FinIpdDetail.Price, FinIpdDetail.DiscTotal, FinIpdDetail.paid, FinIpdMaster.TotalNet,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay ORDER BY FinIpdMaster.Doc_No) = 1 THEN CAST(FinIpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTrim(Str(Day(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinIpdMaster.PrintDateTime) + DATENAME(mi,FinIpdMaster.PrintDateTime) AS Printtime, 
                FinIpdMaster.Paymentcode, FinDetail.Fincode ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinIpdMaster.TotalPay,Payment.PaymentName ORDER BY FinIpdMaster.Doc_No) = 1 THEN Payment.PaymentName ELSE '' END AS PaymentName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinIpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinIpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinIpdMaster.TotalPay) FROM FinIpdMaster  WHERE FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.Docflag <>'F' AND FinIpdMaster.FinDocCode = ('IN5') AND FinIpdMaster.Paymentcode = 'P' AND FinDetail.Fincode <> '218' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate))ELSE NULL END AS TotalPrice
                FROM FinIpdMaster  INNER JOIN PatientData  ON FinIpdMaster.Hn = PatientData.Hn AND FinIpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN  FinIpdDetail  ON FinIpdMaster.Doc_No = FinIpdDetail.Doc_No AND FinIpdMaster.Doc_Yr =FinIpdDetail.Doc_yr AND FinIpdMaster.An = FinIpdDetail.An AND FinIpdMaster.AnYear = FinIpdDetail.AnYear 
                INNER JOIN FinDetail  ON FinIpdDetail.Fincode = FinDetail.Fincode 
                INNER JOIN Payment ON FinIpdMaster.PaymentCode = Payment.PaymentCode
                WHERE FinIpdMaster.DocStatus <> 'V' AND FinIpdMaster.Docflag <>'F' AND FinIpdMaster.FinDocCode = ('IN5') AND FinIpdMaster.Paymentcode = 'P' AND FinDetail.Fincode <> '218' 
                AND (LTrim(STR(DAY(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinIpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinIpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinIpdMaster.DOC_No";  
            }

            using (var con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@Printdate", formatdate);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                con.Open();
                adapter.Fill(dt);
                return dt;
            }
        }

        private MemoryStream ExportpdfIPD(DataTable dt, DateTime Printdate, int Selectoption)
        {
            string[] dataincolumn = null;
            string[] headername = null;
            float[] columnWidth = null;
            string Titlename = string.Empty;
            if (Selectoption == 1)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "AN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "รายการ" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "AN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "PaymentName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "เงินสด";
            }
            else if (Selectoption == 2)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "AN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "ชื่อบริษัท" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "AN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "CompanyName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "ใบแจ้งหนี้";
            }
            else if (Selectoption == 3)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "AN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม"};
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "AN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f};
                Titlename = "ใบแจ้งค่าใช้จ่าย";
            }
            else if (Selectoption == 4)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "AN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "ชื่อบริษัท" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "AN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "CompanyName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "คนไข้ประกันสังคม";
            }
            else if (Selectoption == 5)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "AN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "รายการ" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "AN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "PaymentName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "คนไข้ 30 บาท";
            }
            

            MemoryStream stream = new MemoryStream();
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            writer.CloseStream = false;
            document.Open();

            // ฟ้อนภาษาไทยสำหรับข้อมูลในตาราง
            string fontpath = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont basefont = BaseFont.CreateFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(basefont, 12, Font.NORMAL);

            // ฟ้อนภาษาไทยสำหรับหัวข้อในตาราง
            string fontheader = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont headerbase = BaseFont.CreateFont(fontheader, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font headerfont = new Font(headerbase, 14, Font.BOLD);

            // ฟ้อนภาษาไทยสำหรับหัวข้อใหญ่
            string bigtopic = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont bigtopicbase = BaseFont.CreateFont(bigtopic, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font bigtopicheader = new Font(bigtopicbase, 20, Font.BOLD);

            // ฟ้อนของวันที่และประเภทเอกสาร
            string data = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont database = BaseFont.CreateFont(data, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font datafont = new Font(database, 14, Font.NORMAL);

            // เพื่มชื่อหัวข้อใหญ่
            Paragraph title = new Paragraph("รายงานค่าใช้จ่ายผู้ป่วยใน", bigtopicheader);
            title.Alignment = Element.ALIGN_CENTER;
            title.SpacingAfter = 20f;
            document.Add(title);

            // เพิ่มประเถทเอกสาร
            Paragraph dataparagraph = new Paragraph("ประเภทเอกสาร:" + " " + Titlename, datafont);
            dataparagraph.Alignment = Element.ALIGN_RIGHT;
            document.Add(dataparagraph);

            // เพิ่มวันที่ที่ผู้ใช้เลือก 
            Paragraph dateparagraph = new Paragraph("วันที่" + " " +Printdate.ToString("dd/MM/yyyy"), datafont);
            dateparagraph.Alignment = Element.ALIGN_RIGHT;
            dateparagraph.SpacingAfter = 20f;
            document.Add(dateparagraph);

            // เพิ่มคอลัมล์ตามจำนวนหัวข้อ
            PdfPTable pdfTable = new PdfPTable(dataincolumn.Length);
            pdfTable.WidthPercentage = 100;

            // เพิ่มขนาดแต่ละคอลัมล์ในตาราง
            
            pdfTable.SetWidths(columnWidth);

            foreach (string name in headername)
            {
                PdfPCell cell = new PdfPCell(new Phrase(name, headerfont));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                if (name.Length > 5)
                {
                    cell.NoWrap = false;
                    cell.MinimumHeight = 30f;
                }
                else
                {
                    cell.FixedHeight = 20f;
                }
                pdfTable.AddCell(cell);
            }

            foreach (DataRow row in dt.Rows)
            {
                foreach (string columnname in dataincolumn)
                {
                    //PdfPCell cellValue = new PdfPCell(new Phrase(row[columnname].ToString(), font));
                    string cellValue = row[columnname].ToString(); // เพิ่มหัวข้อ table
                    PdfPCell cell = new PdfPCell(new Phrase(cellValue, font));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;

                    // ถ้ามีข้อมูลหรือคำที่มากกว่า 25 คำ
                    if (cellValue.Length > 25)
                    {
                        cell.NoWrap = false; // อนุญาตให้ขึ้นบรรทัดใหม่
                        cell.MinimumHeight = 30f;
                    }
                    else
                    {
                        cell.FixedHeight = 20f;
                    }
                    pdfTable.AddCell(cell);
                }
            }

            // เพิ่มส่วนล่างของตาราง
            decimal footerdata = 0;

            foreach (DataRow row in dt.Rows)
            {
                if (row["TotalPrice"] != DBNull.Value)
                {
                    footerdata = Convert.ToDecimal(row["TotalPrice"]);
                }
                                
            }
            
            // format การเขียนราคาทั้งหมด
            string[] footerheader = { "TotalPrice : " + " " + footerdata.ToString() };

            PdfPCell footercell = new PdfPCell(new Phrase(footerheader[0], datafont));               
            footercell.Colspan = columnWidth.Length;               
            footercell.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfTable.AddCell(footercell);

            document.Add(pdfTable);    
            document.Close();
              
            stream.Position = 0; // รีเซ็ตตำแหน่งสตรีม
                
            return stream;    
        }

        public ActionResult PdffileIPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = GetexportIPD(Selectoption, Printdate);
            MemoryStream stream = ExportpdfIPD(dt,Printdate,Selectoption);
            return File(stream.ToArray(), "application/pdf", "ReportIPD.pdf");

        }

        private MemoryStream ExportExcelIPD(DataTable dt,int Selectoption, DateTime Printdate)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Data");
                
                int hearderRow = 4;
                string optiontext = "";             
                
                if(Selectoption == 1)
                {
                    optiontext = "เงินสด";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "AN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "รายการ";
                    
                    for(int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1,1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["AN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["PaymentName"];
                    }
                }
                else if(Selectoption == 2)
                {
                    optiontext = "ใบแจ้งหนี้";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "AN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "ชื่อบริษัท";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["AN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["CompanyName"];
                    }
                }
                else if(Selectoption == 3)
                {
                    optiontext = "ใบแจ้งค่ามใช้จ่าย";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "AN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";                  

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["AN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                    }
                }
                else  if(Selectoption == 4)
                {
                    optiontext = "คนไข้ประกันสังคม";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "AN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "ชื่อบริษัท";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["AN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["CompanyName"];
                    }
                }
                else if(Selectoption == 5)
                {
                    optiontext = "คนไข้ 30 บาท";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "AN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "รายการ";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["AN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["PaymentName"];
                    }
                }

                int lastRow = dt.Rows.Count + hearderRow;
                decimal Total = 0;
                foreach (DataRow row in dt.Rows)
                {
                    if (row["TotalPrice"] != DBNull.Value)
                    {
                        Total = Convert.ToDecimal(row["TotalPrice"]);
                    }
                }

                worksheet.Cells[lastRow + 2, 10].Value = "TotalPrice:";
                worksheet.Cells[lastRow + 2,12].Value = Total;

                worksheet.Cells[1, 1].Value = "ประเภทเอกสาร";
                worksheet.Cells[1, 2].Value = optiontext;
                worksheet.Cells[1, 3].Value = "วันที่";
                worksheet.Cells[1, 4].Value = Printdate.ToString("dd/MM/yyyy");

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                return stream;
            }         
        }

        public ActionResult ExcelIPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = GetexportIPD(Selectoption, Printdate);
            MemoryStream stream = ExportExcelIPD(dt, Selectoption, Printdate);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReportIPD.xlsx");
        }




//------------------------------------------------------------------------------------- ข้อมูลสำหรับคนไข้นอก ----------------------------------------------------------------------------//
        private DataTable GetexportOPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = new DataTable();

            // กำหนด format ของวันที่ที่ผู้ใช้กรอกเข้ามา
            var formatdate = $"{Printdate.Day}/{Printdate.Month}/{Printdate.Year}";

            string sql = "";

            // การดึงข้อมูลจาก dropdown
            if (Selectoption == 1)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Doc_No ELSE NULL END AS Doc_No,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Hn ELSE NULL END AS Hn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN YEAR(FinOpdMaster.HnYear) ELSE NULL END AS HnYear,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Vn ELSE NULL END AS Vn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName,
                FinDetail.FinNameT, FinOpdDetail.Price,FinOpdDetail.DiscTotal, FinOpdDetail.Paid, FinOpdMaster.TotalNet,  
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN CAST(FinOpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay,
                LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintdateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinOpdMaster.PrintDateTime) + DATENAME(mi, FinOpdMaster.PrintDateTime) AS PrintTime, 
                FinOpdMaster.PaymentCode, FinDetail.FinCode ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay , Payment.PaymentName ORDER BY FinOpdMaster.Doc_No) = 1 THEN Payment.PaymentName ELSE '' END AS PaymentName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinOpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinOpdMaster.TotalPay) FROM FinOpdMaster WHERE FinOpdMaster.FinDocCode = ('IN1') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'f' AND FinOpdMaster.PaymentCode <> 'P' AND FinOpdMaster.PaymentCode <> 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate))ELSE NULL END AS TotalPrice
                FROM FinOpdMaster 
                INNER JOIN PatientData ON FinOpdMaster.Hn = PatientData.Hn 
                AND FinOpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinOpdDetail ON FinOpdMaster.Doc_No = FinOpdDetail.Doc_No AND FinOpdMaster.Vn = FinOpdDetail.Vn AND FinOpdMaster.VnYear = FinOpdDetail.VnDate AND FinOpdMaster.Doc_Yr = FinOpdDetail.Doc_Yr
                INNER JOIN FinDetail ON FinOpdDetail.FinCode = FinDetail.FinCode
                INNER JOIN Payment ON FinOpdMaster.PaymentCode = Payment.PaymentCode
                WHERE FinOpdMaster.FinDocCode = ('IN1') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'f' AND FinOpdMaster.PaymentCode <> 'P' AND FinOpdMaster.PaymentCode <> 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinOpdMaster.Doc_NO";
            }
            else if (Selectoption == 2)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Doc_No ELSE NULL END AS Doc_No,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Hn ELSE NULL END AS Hn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN YEAR(FinOpdMaster.HnYear) ELSE NULL END AS HnYear,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Vn ELSE NULL END AS Vn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName,
                FinDetail.FinNameT, FinOpdDetail.Price,FinOpdDetail.DiscTotal, FinOpdDetail.Paid, FinOpdMaster.TotalNet, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN CAST(FinOpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay,
                LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintdateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinOpdMaster.PrintDateTime) + DATENAME(mi, FinOpdMaster.PrintDateTime) AS PrintTime, 
                FinOpdMaster.PaymentCode, FinDetail.FinCode ,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay,Companies.CompanyName ORDER BY FinOpdMaster.Doc_No) = 1 THEN Companies.CompanyName ELSE '' END AS CompanyName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinOpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinOpdMaster.TotalPay) FROM FinOpdMaster WHERE FinOpdMaster.FinDocCode = ('IN2') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode <> 'P' AND FinOpdMaster.PaymentCode <> 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                )ELSE NULL END AS TotalPrice
                FROM FinOpdMaster 
                INNER JOIN PatientData ON FinOpdMaster.Hn = PatientData.Hn 
                AND FinOpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinOpdDetail ON FinOpdMaster.Doc_No = FinOpdDetail.Doc_No AND FinOpdMaster.Vn = FinOpdDetail.Vn AND FinOpdMaster.VnYear = FinOpdDetail.VnDate AND FinOpdMaster.Doc_Yr = FinOpdDetail.Doc_Yr
                INNER JOIN FinDetail ON FinOpdDetail.FinCode = FinDetail.FinCode
                INNER JOIN Companies ON FinOpdMaster.CompanyCode = Companies.CompanyCode
                WHERE FinOpdMaster.FinDocCode = ('IN2') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode <> 'P' AND FinOpdMaster.PaymentCode <> 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinOpdMaster.Doc_NO";
            }
            else if (Selectoption == 3)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Doc_No ELSE NULL END AS Doc_No,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Hn ELSE NULL END AS Hn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN YEAR(FinOpdMaster.HnYear) ELSE NULL END AS HnYear,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Vn ELSE NULL END AS Vn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName,
                FinDetail.FinNameT, FinOpdDetail.Price, FinOpdDetail.DiscTotal, FinOpdDetail.Paid, FinOpdMaster.TotalNet, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN CAST(FinOpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay, 
                LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) +'/'+Ltrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinOpdMaster.PrintDateTime) + DATENAME(mi,FinOpdMaster.PrintDateTime) AS Printtime,
                FinOpdMaster.PaymentCode, FinDetail.FinCode, Companies.CompanyName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinOpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinOpdMaster.TotalPay) FROM FinOpdMaster WHERE FinOpdMaster.PHeader = ('ใบแจ้งค่าใช้จ่าย') AND FinOpdMaster.FinDocCode = ('IN3')
                AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode <> 'L'
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) )ELSE NULL END AS TotalPrice
                FROM FinOpdMaster 
                INNER JOIN PatientData ON FinOpdMaster.Hn = PatientData.Hn AND FinOpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinOpdDetail ON FinOpdMaster.Doc_No = FinOpdDetail.Doc_No AND FinOpdMaster.Vn = FinOpdDetail.Vn AND FinOpdMaster.VnYear = FinOpdDetail.VnDate AND FinOpdMaster.Doc_Yr = FinOpdDetail.Doc_Yr 
                INNER JOIN FinDetail ON FinOpdDetail.FinCode = FinDetail.FinCode 
                INNER JOIN Companies ON FinOpdMaster.CompanyCode = Companies.CompanyCode
                WHERE FinOpdMaster.PHeader = ('ใบแจ้งค่าใช้จ่าย') AND FinOpdMaster.FinDocCode = ('IN3')
                AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode <> 'L'
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinOpdMaster.Doc_NO";
            }
            else if (Selectoption == 4)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Doc_No ELSE NULL END AS Doc_No,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Hn ELSE NULL END AS Hn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN YEAR(FinOpdMaster.HnYear) ELSE NULL END AS HnYear,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Vn ELSE NULL END AS Vn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName,
                FinDetail.FinNameT,FinOpdDetail.Price, FinOpdDetail.DiscTotal, FinOpdDetail.Paid, FinOpdMaster.TotalNet, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN CAST(FinOpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay,
                LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) +'/'+Ltrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinOpdMaster.PrintDateTime) + DATENAME(mi,FinOpdMaster.PrintDateTime) AS Printtime, 
                FinOpdMaster.PaymentCode, FinDetail.FinCode, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay,Companies.CompanyName ORDER BY FinOpdMaster.Doc_No) = 1 THEN Companies.CompanyName ELSE '' END AS CompanyName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinOpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinOpdMaster.TotalPay) FROM FinOpdMaster WHERE FinOpdMaster.FinDocCode = ('IN3') 
                AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode = 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) )ELSE NULL END AS TotalPrice
                FROM FinOpdMaster 
                INNER JOIN PatientData ON FinOpdMaster.Hn = PatientData.Hn AND FinOpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinOpdDetail ON FinOpdMaster.Doc_No = FinOpdDetail.Doc_No AND FinOpdMaster.Vn = FinOpdDetail.Vn AND FinOpdMaster.VnYear = FinOpdDetail.VnDate AND FinOpdMaster.Doc_Yr = FinOpdDetail.Doc_Yr 
                INNER JOIN FinDetail ON FinOpdDetail.FinCode = FinDetail.FinCode 
                INNER JOIN Companies ON FinOpdMaster.CompanyCode = Companies.CompanyCode
                WHERE FinOpdMaster.FinDocCode = ('IN3') 
                AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.PaymentCode = 'L' 
                AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinOpdMaster.Doc_NO";
            }
            else if (Selectoption == 5)
            {
                    sql = @"SELECT CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Doc_No ELSE NULL END AS Doc_No,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdDetail.Doc_Yr ELSE NULL END AS Doc_Yr,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Hn ELSE NULL END AS Hn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN YEAR(FinOpdMaster.HnYear) ELSE NULL END AS HnYear,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.Doc_No, FinOpdDetail.Doc_Yr, FinOpdMaster.Hn, YEAR(FinOpdMaster.HnYear) ,FinOpdMaster.Vn ORDER BY FinOpdMaster.Doc_No) = 1 THEN FinOpdMaster.Vn ELSE NULL END AS Vn,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.TitleName ELSE '' END AS TitleName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.FName ELSE '' END AS FName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN PatientData.LName ELSE '' END AS LName,
                FinDetail.FinNameT, FinOpdDetail.Price, FinOpdDetail.DiscTotal, FinOpdDetail.Paid, FinOpdMaster.TotalNet, 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay ORDER BY FinOpdMaster.Doc_No) = 1 THEN CAST(FinOpdMaster.TotalPay AS INT) ELSE NULL END AS TotalPay,
                LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) AS Printdate, DATENAME(hh,FinOpdMaster.PrintDateTime) + DATENAME(mi,FinOPdMaster.PrintDateTime) AS Printtime, 
                FinOpdMaster.PaymentCode,FinDetail.FinCode , 
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY PatientData.TitleName, PatientData.FName, PatientData.LName, FinOpdMaster.TotalPay , Payment.PaymentName ORDER BY FinOpdMaster.Doc_No) = 1 THEN Payment.PaymentName ELSE '' END AS PaymentName,
                CASE WHEN ROW_NUMBER()OVER(PARTITION BY FinOpdMaster.PrintDateTime ORDER BY (SELECT NULL)) = COUNT(*) OVER(PARTITION BY FinOpdMaster.PrintDateTime) THEN
                (SELECT SUM(FinOpdMaster.TotalPay) FROM FinOpdMaster WHERE FinOpdMaster.FinDocCode = ('IN5') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.paymentCode = 'P'
                AND FinDetail.FinCode <> '218' AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate))ELSE NULL END AS TotalPrice
                FROM FinOpdMaster 
                INNER JOIN PatientData ON FinOpdMaster.Hn = PatientData.Hn AND FinOpdMaster.HnYear = PatientData.HnYear 
                INNER JOIN FinOpdDetail ON FinOpdMaster.Doc_No = FinOpdDetail.Doc_No AND FinOpdMaster.Vn = FinOpdDetail.Vn AND FinOpdMaster.VnYear = FinOpdDetail.VnDate AND FinOpdMaster.Doc_Yr = FinOpdDetail.Doc_Yr 
                INNER JOIN FinDetail ON FinOpdDetail.FinCode = FinDetail.FinCode 
                INNER JOIN Payment ON FinOpdMaster.PaymentCode = Payment.PaymentCode
                WHERE FinOpdMaster.FinDocCode = ('IN5') AND FinOpdMaster.DocStatus <> 'V' AND FinOpdMaster.DocFlag <> 'F' AND FinOpdMaster.paymentCode = 'P'
                AND FinDetail.FinCode <> '218' AND (LTrim(STR(DAY(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(MONTH(FinOpdMaster.PrintDateTime))) + '/' + LTrim(STR(YEAR(FinOpdMaster.PrintDateTime))) = @Printdate) 
                ORDER BY FinOpdMaster.Doc_NO";
            }

            using (var con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@Printdate", formatdate);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                con.Open();
                adapter.Fill(dt);

                return dt;
            }

        }
        private MemoryStream ExportpdfOPD(DataTable dt,DateTime Printdate,int Selectoption)
        {
            string[] dataincolumn = null;
            string[] headername = null;
            float[] columnWidth = null;
            string Titlename = string.Empty;
            if (Selectoption == 1)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "VN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "รายการ" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "VN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "PaymentName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "เงินสด";
            }
            else if (Selectoption == 2)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "VN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "ชื่อบริษัท" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "VN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "CompanyName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "ใบแจ้งหนี้";
            }
            else if (Selectoption == 3)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "VN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "VN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f };
                Titlename = "ใบแจ้งค่าใช้จ่าย";
            }
            else if (Selectoption == 4)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "VN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "ชื่อบริษัท" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "VN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "CompanyName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "คนไข้ประกันสังคม";
            }
            else if (Selectoption == 5)
            {
                headername = new string[] { "เลขที่ใบเสร็จ", "ปี", "HN", "HnYear", "VN", "คำนำหน้า", "ชื่อ", "นามสกุล", "รายการ", "จำนวนเงิน", "ส่วนลด", "รวม", "รายการ" };
                dataincolumn = new string[] { "Doc_No", "Doc_Yr", "HN", "HNYear", "VN", "Titlename", "FName", "LName", "FinNameT", "Price", "Disctotal", "TotalPay", "PaymentName" };
                columnWidth = new float[] { 1.1f, 1f, 1.2f, 1.2f, 1.2f, 1.2f, 1.5f, 1.2f, 3.5f, 1.5f, 1f, 1.5f, 2.5f };
                Titlename = "คนไข้ 30 บาท";
            }

            MemoryStream stream = new MemoryStream();
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            writer.CloseStream = false;
            document.Open();
            // ฟ้อนภาษาไทยสำหรับข้อมูลในตาราง
            string fontpath = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont basefont = BaseFont.CreateFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(basefont, 12, Font.NORMAL);

            // ฟ้อนภาษาไทยสำหรับหัวข้อในตาราง
            string fontheader = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont headerbase = BaseFont.CreateFont(fontheader, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font headerfont = new Font(headerbase, 14, Font.BOLD);

            // ฟ้อนภาษาไทยสำหรับหัวข้อใหญ่
            string bigtopic = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont bigtopicbase = BaseFont.CreateFont(bigtopic, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font bigtopicheader = new Font(bigtopicbase, 20, Font.BOLD);

            // ฟ้อนของวันที่และประเภทเอกสาร
            string data = Path.Combine(Directory.GetCurrentDirectory(), "Font", "THSarabunNew.ttf");
            BaseFont database = BaseFont.CreateFont(data, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font datafont = new Font(database, 14, Font.NORMAL);

            // เพื่มชื่อหัวข้อใหญ่
            Paragraph title = new Paragraph("รายงานค่าใช้จ่ายผู้ป่วยนอก", bigtopicheader);
            title.Alignment = Element.ALIGN_CENTER;
            title.SpacingAfter = 20f;
            document.Add(title);

            // เพิ่มประเถทเอกสาร
            Paragraph dataparagraph = new Paragraph("ประเภทเอกสาร:" + " " + Titlename, datafont);
            dataparagraph.Alignment = Element.ALIGN_RIGHT;
            document.Add(dataparagraph);

            // เพิ่มวันที่ที่ผู้ใช้เลือก 
            Paragraph dateparagraph = new Paragraph("วันที่" + " " + Printdate.ToString("dd/MM/yyyy"), datafont);
            dateparagraph.Alignment = Element.ALIGN_RIGHT;
            dateparagraph.SpacingAfter = 20f;
            document.Add(dateparagraph);

            // เพิ่มคอลัมล์ตามจำนวนหัวข้อ
            PdfPTable pdfTable = new PdfPTable(dataincolumn.Length);
            pdfTable.WidthPercentage = 100;

            // เพิ่มขนาดแต่ละคอลัมล์ในตาราง
            pdfTable.SetWidths(columnWidth);

            foreach (string name in headername)
            {
                PdfPCell cell = new PdfPCell(new Phrase(name, headerfont));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                if (name.Length > 5)
                {
                    cell.NoWrap = false;
                    cell.MinimumHeight = 30f;
                }
                else
                {
                    cell.FixedHeight = 20f;
                }
                pdfTable.AddCell(cell);
            }

            foreach (DataRow row in dt.Rows)
            {
                foreach (string columnname in dataincolumn)
                {
                    string cellValue = row[columnname].ToString(); // เพิ่มหัวข้อ table
                    PdfPCell cell = new PdfPCell(new Phrase(cellValue, font));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;

                    // ถ้ามีข้อมูลหรือคำที่มากกว่า 25 คำ
                    if (cellValue.Length > 25)
                    {
                        cell.NoWrap = false;
                        cell.MinimumHeight = 30f;
                    }
                    else
                    {
                        cell.FixedHeight = 20f;
                    }
                    pdfTable.AddCell(cell);
                }
            }
            // เพิ่มส่วนล่างของตาราง
            decimal footerdata = 0;

            foreach (DataRow row in dt.Rows)
            {
                if (row["TotalPrice"] != DBNull.Value)
                {
                    footerdata = Convert.ToDecimal(row["TotalPrice"]);
                }

            }

            // format การเขียนราคาทั้งหมด
            string[] footerheader = { "TotalPrice : " + " " + footerdata.ToString() };

            PdfPCell footercell = new PdfPCell(new Phrase(footerheader[0], datafont));
            footercell.Colspan = columnWidth.Length;
            footercell.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfTable.AddCell(footercell);

            document.Add(pdfTable);
            document.Close();

            stream.Position = 0; // รีเซ็ตตำแหน่งสตรีม
            return stream;
        }

        public ActionResult PdffileOPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = GetexportOPD(Selectoption, Printdate);
            MemoryStream stream = ExportpdfOPD(dt,Printdate,Selectoption);
            return File(stream.ToArray(), "application/pdf", "ReportOPD.pdf");
        }

        private MemoryStream ExportExcelOPD(DataTable dt, int Selectoption, DateTime Printdate)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Data");

                int hearderRow = 4;
                string optiontext = "";

                if (Selectoption == 1)
                {
                    optiontext = "เงินสด";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "VN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "รายการ";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["VN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["PaymentName"];
                    }
                }
                else if (Selectoption == 2)
                {
                    optiontext = "ใบแจ้งหนี้";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "VN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "ชื่อบริษัท";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["VN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["CompanyName"];
                    }
                }
                else if (Selectoption == 3)
                {
                    optiontext = "ใบแจ้งค่าใช้จ่าย";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "VN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["VN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                    }
                }
                else if (Selectoption == 4)
                {
                    optiontext = "คนไข้ประกันสังคม";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "VN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "ชื่อบริษัท";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["VN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["CompanyName"];
                    }
                }
                else if (Selectoption == 5)
                {
                    optiontext = "คนไข้ 30 บาท";
                    worksheet.Cells[hearderRow, 1].Value = "เลขที่ใบเสร็จ";
                    worksheet.Cells[hearderRow, 2].Value = "ปี";
                    worksheet.Cells[hearderRow, 3].Value = "HN";
                    worksheet.Cells[hearderRow, 4].Value = "HNYear";
                    worksheet.Cells[hearderRow, 5].Value = "VN";
                    worksheet.Cells[hearderRow, 6].Value = "คำนำหน้า";
                    worksheet.Cells[hearderRow, 7].Value = "ชื่อ";
                    worksheet.Cells[hearderRow, 8].Value = "นามสกุล";
                    worksheet.Cells[hearderRow, 9].Value = "รายการ";
                    worksheet.Cells[hearderRow, 10].Value = "จำนวนเงิน";
                    worksheet.Cells[hearderRow, 11].Value = "ส่วนลด";
                    worksheet.Cells[hearderRow, 12].Value = "รวม";
                    worksheet.Cells[hearderRow, 13].Value = "รายการ";

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        worksheet.Cells[hearderRow + row + 1, 1].Value = dt.Rows[row]["Doc_No"];
                        worksheet.Cells[hearderRow + row + 1, 2].Value = dt.Rows[row]["Doc_Yr"];
                        worksheet.Cells[hearderRow + row + 1, 3].Value = dt.Rows[row]["HN"];
                        worksheet.Cells[hearderRow + row + 1, 4].Value = dt.Rows[row]["HNYear"];
                        worksheet.Cells[hearderRow + row + 1, 5].Value = dt.Rows[row]["VN"];
                        worksheet.Cells[hearderRow + row + 1, 6].Value = dt.Rows[row]["TitleName"];
                        worksheet.Cells[hearderRow + row + 1, 7].Value = dt.Rows[row]["FName"];
                        worksheet.Cells[hearderRow + row + 1, 8].Value = dt.Rows[row]["LName"];
                        worksheet.Cells[hearderRow + row + 1, 9].Value = dt.Rows[row]["FinNameT"];
                        worksheet.Cells[hearderRow + row + 1, 10].Value = dt.Rows[row]["Price"];
                        worksheet.Cells[hearderRow + row + 1, 11].Value = dt.Rows[row]["Disctotal"];
                        worksheet.Cells[hearderRow + row + 1, 12].Value = dt.Rows[row]["TotalPay"];
                        worksheet.Cells[hearderRow + row + 1, 13].Value = dt.Rows[row]["PaymentName"];
                    }
                }

                int lastRow = dt.Rows.Count + hearderRow;
                decimal Total = 0;
                foreach (DataRow row in dt.Rows)
                {
                    if (row["TotalPrice"] != DBNull.Value)
                    {
                        Total = Convert.ToDecimal(row["TotalPrice"]);
                    }
                }

                worksheet.Cells[lastRow + 2, 10].Value = "TotalPrice:";
                worksheet.Cells[lastRow + 2, 12].Value = Total;

                worksheet.Cells[1, 1].Value = "ประเภทเอกสาร";
                worksheet.Cells[1, 2].Value = optiontext;
                worksheet.Cells[1, 3].Value = "วันที่";
                worksheet.Cells[1, 4].Value = Printdate.ToString("dd/MM/yyyy");

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                return stream;
            }
        }

        public ActionResult ExcelOPD(int Selectoption, DateTime Printdate)
        {
            DataTable dt = GetexportOPD(Selectoption, Printdate);
            MemoryStream stream = ExportExcelOPD(dt, Selectoption, Printdate);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReportOPD.xlsx");
        }

    }
}
