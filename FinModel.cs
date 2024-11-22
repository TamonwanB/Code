using Ruamchai.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Ruamchai.Data;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;


namespace Ruamchai.Models
{
    public class FinModel
    {
        public int Doc_No { get; set; }
        public int Doc_Yr { get; set; }
        public int Hn { get; set; }
        public int HnYear { get; set; }
        public int An { get; set; }
        public string TitleName { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
        public string FinNameT { get; set; }
        public double Price { get; set; }
        public double DiscTotal { get; set; }
        public double Paid { get; set; }
        public double TotalNet { get; set; }
        public double TotalPay { get; set; }
        public DateTime Printdate { get; set; }
        public TimeSpan PrinttimeStart { get; set; }
        public TimeSpan PrinttimeEnd {get; set;}
        public string PaymentCode { get; set; }
        public string FinCode { get; set; }
    }

    public class TimeRangeModel
    {
        //[RegularExpression(@"^(00|01|02|03|04|05|06|07|08|09|1[0-9]|2[0-3]:[0-5][0-9]$",
        //    ErrorMessage = "กรุณาใส่เวลาให้ถูกต้อง")]
        //[DataType(DataType.Time)]
        public string StartTime { get; set; }

        //[RegularExpression(@"^(00|01|02|03|04|05|06|07|08|09|1[0-9]|2[0-3]:[0-5][0-9]$",
        //    ErrorMessage = "กรุณาใส่เวลาให้ถูกต้อง")]
        //[DataType(DataType.Time)]
        public string EndTime { get; set; }

        //[DataType(DataType.Date)]
        //[DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime SelectDate { get; set; }
    }

    public class CombineModel
    {
        public FinModel FinModel { get; set; }
        public TimeRangeModel TimeRangeModel { get; set; }
        public optionModel optionModel { get; set; }
        public List<FinModel> Finmodels { get; set; }
        public List<TimeRangeModel> TimeRangeModels { get; set; }
        public List<optionModel> optionModels { get; set; }
    }

    public class optionModel
    {
        public string Selectoption { get; set; }
        public List<SelectListItem> options { get; set; }
        public List<FinModel> code { get; set; }  
    }
}
