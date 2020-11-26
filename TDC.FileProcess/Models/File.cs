namespace TDC.FileProcess.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("Files")]
    public partial class Files
    {
        public int Id { get; set; }

        [StringLength(255)]
        public string Code { get; set; }

        [StringLength(255)]
        public string FullName { get; set; }

        [StringLength(255)]
        public string Department { get; set; }

        public DateTime? DateWorking { get; set; }

        [StringLength(255)]
        public string CheckIn { get; set; }

        public string CheckOut { get; set; }
    }
}
