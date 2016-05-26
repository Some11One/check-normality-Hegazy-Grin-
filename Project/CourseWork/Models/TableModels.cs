using System.ComponentModel.DataAnnotations;
using System.Data.Entity;

namespace CourseWork.Models
{
    public class TableModels 
    {
        [Key]
        public int StudID { get; set; } // Student's ID
        public int RawScore { get; set; } // Student's score
        public double Z { get; set; } // Z score
        public double T { get; set; } // T score
        public double Perc { get; set; } // Percentile of a student
        static public TableModels[] tableArray { get; set; }
    }

    public class TableModelsDbContext : DbContext
    {
        public DbSet<TableModels> Students { get; set; }
    }
}