using System.ComponentModel.DataAnnotations;
using System.Web;
using System.Data.Entity;
using System;

namespace CourseWork.Models
{
    public class UploadModels
    {
        [Required]
        public HttpPostedFileBase File { get; set; } // File that will be uploaded

        static public string[,] Data { get; set; } // Table data from uploaded file

        static public double[] Y { get; set; } // Data for Z scalе

        static public double S { get; set; } // Standart Deviation

        static public double X { get; set; } // Mean Value

        static public double[] Perc { get; set; } // Percentiles

        static public int[] Answers { get; set; } // Answers of each student
    }
}