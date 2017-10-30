using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace samples
{
    public class Report
    {
        [Display(Name = "城市")]
        public string City { get; set; }

        [Display(Name = "楼盘")]
        public string Building { get; set; }

        [Display(Name = "区域")]
        public string Area { get; set; }

        [Display(Name = "成交时间")]
        public DateTime HandleTime { get; set; }

        [Display(Name = "经纪人")]
        public string Broker { get; set; }

        [Display(Name = "客户")]
        public string Customer { get; set; }

        [Display(Name = "房源")]
        public string Room { get; set; }

        [Display(Name = "佣金(元)")]
        public decimal Brokerage { get; set; }

        [Display(Name = "收益(元)")]
        public decimal Profits { get; set; }

        public Customer CustomerObj { get; set; }
    }

    [ComplexType]
    public class Customer
    {
        [Display(Name = "ID")]
        public long Id { get; set; }

        [Display(Name = "LastName")]
        public string LastName { get; set; }

        [Display(Name = "FirstName")]
        public string FirstName { get; set; }
    }
}