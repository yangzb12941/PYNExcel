using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace PYNExcel
{
    public partial class pynForm : Form
    {
        private ReadExcelTool readExcelTool = null;//读excel
        private List<Ratio> ratioList = new List<Ratio>(8);//存放补贴比率
        private HashSet<string> comCustomsType = new HashSet<string>(8);//公司-清关类型
        private HashSet<string> dataGridViewValue = new HashSet<string>(8);//公司-清关类型
        private Dictionary<string,int> companyNameDic = new Dictionary<string, int>(8);//checkListBox 当前已有的公司名称
        public pynForm()
        {
            InitializeComponent();
        }

        private void checkFileButton_Click(object sender, EventArgs e)
        {
            string file = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;      //该值确定是否可以选择多个文件
            dialog.Title = "请选择Excel文件";     //弹窗的标题
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); //默认打开的桌面的位置
            dialog.Filter = "MicroSoft Excel文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";       //筛选文件
            dialog.ShowHelp = true;     //是否显示“帮助”按钮

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                file = dialog.FileName;
            }
            else 
            {
                MessageBox.Show("请选择excel文件!");
                return;
            }
            this.fileNameTextBox.Text = file;
            readExcelTool = new ReadExcelTool(file);
            bindSheetNameToCheckedSheetListBox();
        }

        private void bindSheetNameToCheckedSheetListBox()
        {
            Workbook workbook = readExcelTool.getWorkbook();
            List<string> sheetNameList = new List<string>(workbook.Worksheets.Count);
            for (int i = 1; i <= workbook.Worksheets.Count; i++)
            {
                //取得第一个工作薄
                Worksheet ws = (Worksheet)workbook.Worksheets.get_Item(i);
                sheetNameList.Add(ws.Name);
            }
            checkedSheetListBoxAddItem(sheetNameList);
        }

        private void checkedSheetListBoxAddItem(List<string> sheetNameList) 
        {
            this.checkedSheetListBox.Items.Clear();

            for (int i = 0; i < sheetNameList.Count; i++) 
            {
                this.checkedSheetListBox.Items.Add(sheetNameList[i], false);
            }
        }

        private void sheetTrueButton_Click(object sender, EventArgs e)
        {
           int checkCount = this.checkedSheetListBox.CheckedItems.Count;
            List<string> sheetNameList = new List<string>(checkCount);
            for (int i = 0; i < checkCount; i++) 
            {
                sheetNameList.Add(this.checkedSheetListBox.CheckedItems[i].ToString());
            }
            goodsBindToCheckedGoodsListBox(sheetNameList);
        }

        private void goodsBindToCheckedGoodsListBox(List<string> sheetNameList) 
        {
            this.checkedGoodsListBox.Items.Clear();
            //清楚列表
            while (this.dataGridView.Rows.Count > 1)
            {
                this.dataGridView.Rows.RemoveAt(1);
            }
            //公司-产品
            HashSet<string> goodsNames = new HashSet<string>(32);
            for (int i = 0; i < sheetNameList.Count; i++) 
            {
                Worksheet worksheet = readExcelTool.getWorksheet(sheetNameList[i]);
                //遍历行，获取对应行某些列的数据
                //取得总记录行数   (包括标题列)
                int rowsint = worksheet.UsedRange.Rows.Count; //得到行数

                //默认从第3行开始解析数据
                for (int rowsCount = 3; rowsCount < rowsint; rowsCount++)
                {
                    string cellAValue = ((Range)worksheet.Cells[rowsCount,"A"]).Text.ToString();
                    //读取第A\H列
                    string cellHValue = ((Range)worksheet.Cells[rowsCount, "H"]).Text.ToString();
                    //清关类型
                    string cellBOValue = ((Range)worksheet.Cells[rowsCount, "BO"]).Text.ToString();
                    if (String.IsNullOrWhiteSpace(cellHValue)) 
                    {
                        continue;
                    }
                    //公司-品种
                    goodsNames.Add(cellAValue + "-" + cellHValue);
                    string keyType = cellAValue + "-" + cellBOValue;
                    if (!comCustomsType.Contains(keyType))
                    {
                        //公司-清关类型
                        comCustomsType.Add(keyType);
                        Ratio ratio = new Ratio();
                        ratio.CompanyName = cellAValue;
                        ratio.CustomstType = cellBOValue;
                        //被勾选
                        ratioList.Add(ratio);//存放补贴比率
                    }
                }
            }

            foreach (string goodsName in goodsNames)
            {
                this.checkedGoodsListBox.Items.Add(goodsName, false);
            }
        }

        private void handleButton_Click(object sender, EventArgs e)
        {
            //勾选了物料之后，根据物料的公司抬头、物料名称，获取对应的提单金额
            int checkCount = this.checkedGoodsListBox.CheckedItems.Count;
            List<CompanyGoods> companyGoodsList = new List<CompanyGoods>(64);
            for (int i = 0; i < checkCount; i++)
            {
                //获取已选择的公司-物料
                string comGoods = this.checkedGoodsListBox.CheckedItems[i].ToString();
                //公司抬头名称
                string companyName = comGoods.Split('-')[0];
                //国别&品种
                string goodName = comGoods.Split('-')[1];
                int checkSheetCount = this.checkedSheetListBox.CheckedItems.Count;

                for (int iSheet = 0; iSheet < checkSheetCount; iSheet++)
                {
                    string sheetName = this.checkedSheetListBox.CheckedItems[iSheet].ToString();
                    List<CompanyGoods> companyGoods = getCompanyGoods(companyName, goodName, sheetName);
                    if (null != companyGoods && companyGoods.Any()) 
                    {
                        companyGoodsList.AddRange(companyGoods);
                    }
                }
            }
            calGoodsUSD(companyGoodsList);
        }

        private List<CompanyGoods> getCompanyGoods(string companyName, string goodName,string sheetName) 
        {
            List < CompanyGoods > companyGoods =new List<CompanyGoods>(16);
            Worksheet worksheet = readExcelTool.getWorksheet(sheetName);
            //遍历行，获取对应行某些列的数据
            //取得总记录行数   (包括标题列)
            int rowsint = worksheet.UsedRange.Rows.Count; //得到行数
            Boolean isCleared = sheetName.IndexOf("已结关") >= 0 ? true:false;
            //默认从第3行开始解析数据
            for (int rowsCount = 3; rowsCount <= rowsint; rowsCount++)
            {
                string cellAValue = ((Range)worksheet.Cells[rowsCount, "A"]).Text.ToString();
                //读取第A\H列
                string cellHValue = ((Range)worksheet.Cells[rowsCount, "H"]).Text.ToString();

                //读取第BO列-清关类型
                string cellBOValue = ((Range)worksheet.Cells[rowsCount, "BO"]).Text.ToString();
                if (cellAValue.Equals(companyName) && cellHValue.Equals(goodName))
                {
                    //读取第R列提单金额(USD)
                    string cellRValue = ((Range)worksheet.Cells[rowsCount, "R"]).Text.ToString();
                    CompanyGoods companyGood = new CompanyGoods();
                    companyGood.CompanyName = companyName;
                    companyGood.GoodsName = goodName;
                    companyGood.CustomstType = cellBOValue;
                    companyGood.Usd = Double.Parse(cellRValue);
                    companyGood.IsCleared = isCleared;
                    companyGoods.Add(companyGood);
                }
            }
            return companyGoods;
        }

        private void calGoodsUSD(List<CompanyGoods> companyGoodsList)
        {
            Dictionary<string,StatisticalData> statisticalDatas = new Dictionary<string,StatisticalData>(8);
            foreach (CompanyGoods cGoodsItem in companyGoodsList)
            {
                string key = cGoodsItem.CompanyName + "-" + cGoodsItem.CustomstType + "-" + cGoodsItem.GoodsName;
                if (statisticalDatas.ContainsKey(key))
                {
                    StatisticalData sData = statisticalDatas[key];
                    if (cGoodsItem.IsCleared)
                    {
                        sData.ClearedQuota += cGoodsItem.Usd;
                    }
                    else
                    {
                        sData.CustomsClearance += cGoodsItem.Usd;
                    }
                    sData.CustomsTotal += cGoodsItem.Usd;
                    sData.IncrementalSubsidy = 0d;
                }
                else 
                {
                    StatisticalData sData = new StatisticalData();
                    sData.CompanyName = cGoodsItem.CompanyName;
                    sData.GoodsName = cGoodsItem.GoodsName;
                    sData.CustomstType = cGoodsItem.CustomstType;
                    if (cGoodsItem.IsCleared)
                    {
                        sData.ClearedQuota = cGoodsItem.Usd;
                        sData.CustomsClearance = 0d;
                    }
                    else 
                    {
                        sData.CustomsClearance = cGoodsItem.Usd;
                        sData.ClearedQuota = 0d;
                    }
                    sData.CustomsTotal += cGoodsItem.Usd;
                    sData.IncrementalSubsidy = 0d;
                    Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(cGoodsItem.CompanyName) && item.CustomstType.Equals(cGoodsItem.CustomstType));//获取补贴比率对象
                    sData.Ratio = ratio.RatioValue;
                    statisticalDatas.Add(key, sData);
                }
            }
            calValue(statisticalDatas);
        }

        private void calValue(Dictionary<string, StatisticalData> statisticalDatas)
        {

            List<StatisticalData> statisticalDataList = new List<StatisticalData>(statisticalDatas.Count);

            //已结关额度
            double clearedQuotaAll = 0.0d;
            //清关中额度
            double customsClearanceAll = 0.0d;
            //总额度
            double customsTotalAll = 0.0d;
            //增量补贴
            double incrementalSubsidyAll = 0.0d;

            foreach (var entry in statisticalDatas)
            {
                string key = entry.Key;
                StatisticalData aData = entry.Value;
                aData.IncrementalSubsidy = aData.CustomsTotal * aData.Ratio;

                clearedQuotaAll += aData.ClearedQuota;
                customsClearanceAll += aData.CustomsClearance;
                customsTotalAll += aData.CustomsTotal;
                incrementalSubsidyAll += aData.IncrementalSubsidy;
                statisticalDataList.Add(aData);
            }
            StatisticalData statisticalDataEnd = new StatisticalData();
            statisticalDataEnd.CompanyName = "合计";
            statisticalDataEnd.GoodsName = "/";
            statisticalDataEnd.CustomstType = "/";
            statisticalDataEnd.ClearedQuota = clearedQuotaAll;//已结关额度
            statisticalDataEnd.CustomsClearance = customsClearanceAll;//清关中额度
            statisticalDataEnd.CustomsTotal = customsTotalAll;//总额度
            statisticalDataEnd.IncrementalSubsidy = incrementalSubsidyAll;//增量补贴
            statisticalDataList.Add(statisticalDataEnd);
            writeExcel(statisticalDataList);
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //编辑补贴率
            DataGridViewRow rowItem = this.dataGridView.Rows[e.RowIndex];
            
            if (null != rowItem.Cells[2].Value)
            {
                string ratioValue = rowItem.Cells[2].Value.ToString();
                string dataCompany = rowItem.Cells[0].Value.ToString();
                string customstType = rowItem.Cells[1].Value.ToString();
                Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(dataCompany) && item.CustomstType.Equals(customstType));//获取补贴比率对象
                ratio.RatioValue = Double.Parse(ratioValue);
            }
            else 
            {
                string dataCompany = rowItem.Cells[0].Value.ToString();
                string customstType = rowItem.Cells[1].Value.ToString();
                Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(dataCompany) && item.CustomstType.Equals(customstType));//获取补贴比率对象
                ratio.RatioValue = 1.0d;
            }
        }

        private void writeExcel(List<StatisticalData> statisticalDatas) 
        {
            WriteExcelTool writeExcelTool = new WriteExcelTool(statisticalDatas);
            string fileName = writeExcelTool.writeExcel();
            if (String.IsNullOrEmpty(fileName))
            {
                MessageBox.Show("Excel 创建失败!");
            }
            else 
            {
                MessageBox.Show("Excel 创建成功:" + fileName);
            }
        }

        private void checkedGoodsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckedListBox checkedListBox = (CheckedListBox)sender;
            int index = checkedListBox.SelectedIndex;//idx就是当前的选中项序号
            string value = checkedListBox.GetItemText(checkedListBox.Items[index]);
            if (checkedListBox.GetItemChecked(index))
            {
                string companyName = value.Split('-')[0];
                if (companyNameDic.ContainsKey(companyName))
                {
                    companyNameDic[companyName]++;
                }
                else 
                {
                    companyNameDic.Add(companyName, 1);
                }
                foreach (Ratio ratio in ratioList)
                {
                    if (companyName.Equals(ratio.CompanyName))
                    {
                        string key = ratio.CompanyName + "-" + ratio.CustomstType;
                        if (dataGridViewValue.Contains(key)) 
                        {
                            continue;
                        }
                        dataGridViewValue.Add(key);
                        int dataIndex = this.dataGridView.Rows.Add();
                        this.dataGridView.Rows[dataIndex].Cells[0].Value = ratio.CompanyName;//公司抬头
                        this.dataGridView.Rows[dataIndex].Cells[1].Value = ratio.CustomstType;//清关类型
                    }
                    else 
                    {
                        continue;
                    }
                }
            }
            else
            {
                //不被勾选
                string companyName = value.Split('-')[0];
                if (companyNameDic.ContainsKey(companyName))
                {
                   int timesCount = --companyNameDic[companyName];
                    if (timesCount <= 0)
                    {
                        companyNameDic.Remove(companyName);
                        int rowount = dataGridView.Rows.Count;//得到总行数
                        List<DataGridViewRow> delRows = new List<DataGridViewRow>(rowount);
                        for (int i =0; i < rowount; i++)
                        {
                            string companyNameDGV = dataGridView.Rows[i].Cells[0].Value.ToString();
                            if (companyNameDGV.Equals(companyName)) 
                            {
                                var rowItem = dataGridView.Rows[i];
                                delRows.Add(rowItem);
                            }
                        }

                        foreach (DataGridViewRow dgvr in delRows)
                        {
                            dataGridView.Rows.Remove(dgvr);
                        }

                        dataGridViewValue.RemoveWhere(s => { return s.Split('-')[0].Equals(companyName);});
                    }
                }
            }
        }
    }
}
