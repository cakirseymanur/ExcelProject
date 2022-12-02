using ExcelProject.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelProject.Controllers
{
    public class ExhangeRateController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ExchangeRateExcel()
        {
            string url = "http://www.tcmb.gov.tr/kurlar/";
            string date = DateTime.Now.AddDays(-1).ToShortDateString();

            ExcelPackage excelPackage = new ExcelPackage();
            var workSheet = excelPackage.Workbook.Worksheets.Add("Kur Bilgileri");
            workSheet.Cells[1, 1].Value = "Tarih";
            workSheet.Cells[1, 2].Value = "Kur";
            workSheet.Cells[1, 3].Value = "EURO/USD";

            List<ExchangeRate> kurUsdListesi = new List<ExchangeRate>();
            List<ExchangeRate> kurEuroListesi = new List<ExchangeRate>();
            for (int i = -1; i > -31; i--)
            {

                try
                {

                    ExchangeRate usdList = new ExchangeRate();
                    ExchangeRate euroList  = new ExchangeRate();
                    url = DateConvertXml(date);
                    var xmldoc = new XmlDocument();
                    xmldoc.Load(url);
                    
                    euroList.Kur = Convert.ToDouble(xmldoc.SelectSingleNode("Tarih_Date/Currency [@Kod='EUR']/BanknoteSelling").InnerXml.Replace('.', ','));
                    euroList.Date = date;

                    usdList.Kur = Convert.ToDouble(xmldoc.SelectSingleNode("Tarih_Date/Currency [@Kod='USD']/BanknoteSelling").InnerXml.Replace('.', ','));
                    usdList.Date = date;

                    kurUsdListesi.Add(usdList);
                    kurEuroListesi.Add(euroList);
                    date = DateTime.Now.AddDays(i - 1).ToShortDateString();
                }
                catch (Exception)
                {
                    i--;
                    date = DateTime.Now.AddDays(i - 1).ToShortDateString();
                    continue;

                }

            }

            List<ExchangeRate> usdListDesc = kurUsdListesi.OrderByDescending(x => x.Kur).ToList();
            List<ExchangeRate> euroListDesc = kurEuroListesi.OrderByDescending(x => x.Kur).ToList();
            int row = 2;

            for (int i = 0; i < 5; i++)
            {
                workSheet.Cells[row, 1].Value = usdListDesc[i].Date;
                workSheet.Cells[row, 2].Value = usdListDesc[i].Kur;
                workSheet.Cells[row, 3].Value = "$ USD";
                row++;
            }
            for (int i = 0; i < 5; i++)
            {
                workSheet.Cells[row, 1].Value = euroListDesc[i].Date; 
                workSheet.Cells[row, 2].Value = euroListDesc[i].Kur; 
                workSheet.Cells[row, 3].Value = "€ EURO"; 
                row++;
            }
            var bytes = excelPackage.GetAsByteArray();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Kurlar.xlsx");
 
        }
        public string DateConvertXml(string date)
        {
            string url = "http://www.tcmb.gov.tr/kurlar/";
            List<string> date1 = date.Split('.').ToList();
            if (date1[0].Length <= 1)
            {
                date1[0] = "0" + date1[0];
            }
            if (date1[1].Length <= 1)
            {
                date1[1] = "0" + date1[1];
            }
            string yearMonth = date1[2] + date1[1];
            string toDay = date1[0] + date1[1] + date1[2];
            url = url + yearMonth + "/" + toDay + ".xml";
            return url;
        }

    }
}
