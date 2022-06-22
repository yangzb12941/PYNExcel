using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PYNExcel
{
    internal class StatisticalData
    {
        //产品
        private string goodsName;

        //公司抬头
        private string companyName;

        //已结关额度
        private double clearedQuota;

        //清关中额度
        private double customsClearance;

        //总额度
        private double customsTotal;

        //增量补贴
        private double incrementalSubsidy;

        //补贴率
        private double ratio;

        public string GoodsName { get => goodsName; set => goodsName = value; }
        public string CompanyName { get => companyName; set => companyName = value; }
        public double ClearedQuota { get => clearedQuota; set => clearedQuota = value; }
        public double CustomsClearance { get => customsClearance; set => customsClearance = value; }
        public double CustomsTotal { get => customsTotal; set => customsTotal = value; }
        public double IncrementalSubsidy { get => incrementalSubsidy; set => incrementalSubsidy = value; }
        public double Ratio { get => ratio; set => ratio = value; }
    }
}
