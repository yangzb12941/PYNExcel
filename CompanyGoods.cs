using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PYNExcel
{
    internal class CompanyGoods
    {
        //公司抬头
        private string companyName;
        //国别&品种
        private string goodsName;
        //提单金额
        private double usd;
        //是否已关结
        private Boolean isCleared;

        public string CompanyName { get => companyName; set => companyName = value; }
        public string GoodsName { get => goodsName; set => goodsName = value; }
        public double Usd { get => usd; set => usd = value; }
        public bool IsCleared { get => isCleared; set => isCleared = value; }
    }
}
