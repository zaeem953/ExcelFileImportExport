using System.ComponentModel.DataAnnotations;

namespace ExcelFileImportExport.Models
{
    public class ExcelCustomer
    {
        [Key]
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Gender { get; set; }
        public string Country    { get; set; }
        public int Age { get; set; }
    }
}
