using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PYNExcel
{
    internal class Ratio
    {
        private string companyName;
        private string goodsName;
        private double ratioValue;

        public string CompanyName { get => companyName; set => companyName = value; }
        public string GoodsName { get => goodsName; set => goodsName = value; }
        public double RatioValue { get => ratioValue; set => ratioValue = value; }
    }
}
