namespace TDC.FileProcess.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("Files")]
    public partial class Files
    {
        public int Id { get; set; }

        [StringLength(50)]
        public string Code { get; set; }

        [StringLength(50)]
        public string FullName { get; set; }

        [StringLength(100)]
        public string Department { get; set; }

        public DateTime? DateWorking { get; set; }

        [StringLength(50)]
        public string CheckIn { get; set; }

        [StringLength(50)]
        public string CheckOut { get; set; }
    }
}
