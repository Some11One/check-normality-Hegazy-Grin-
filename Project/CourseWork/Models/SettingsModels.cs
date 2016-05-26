
namespace CourseWork.Models
{
    /// <summary>
    /// Class to store data about user's settings of application
    /// </summary>
    public class SettingsModels
    {
        static public bool Abs = true; // Pure answers
        static public bool RawScore = true; // How many students answered on questions
        static public bool Perc = true; // Percentiles
        static public bool Ckvar25 = true; // 25 bottom perc
        static public bool Ckvar75 = true; // 25 top perc
        static public bool Z = true; // Z-score
        static public bool T = true; // T-scpre
        static public bool Table = false; // Table of each student's results

        /// <summary>
        /// How to sort table in "Results" page
        /// </summary>
        /// <remarks> True - by score, False - by id of a student
        /// </remarks>
        static public bool Sort = true; 
    }
}