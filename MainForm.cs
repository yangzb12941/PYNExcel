using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.ListBox;

namespace PYNExcel
{
    public partial class pynForm : Form
    {
        private ReadExcelTool readExcelTool = null;//读excel
        private WriteExcelTool writeExcelTool = null;//写excel
        private List<Ratio> ratioList = new List<Ratio>(8);//存放补贴比率
        private HashSet<String> dataSet = new HashSet<string>(8);//加入dataGridView中的数据过滤器

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
            HashSet<string> goodsNames = new HashSet<string>();
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
                    if (String.IsNullOrWhiteSpace(cellHValue)) 
                    {
                        continue;
                    }
                    goodsNames.Add(cellAValue + "-" + cellHValue);
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
            Dictionary<String,List < CompanyGoods >> keyValues = new Dictionary<String, List < CompanyGoods>>(checkCount);
            for (int i = 0; i < checkCount; i++)
            {
                List<CompanyGoods> companyGoodsList = new List<CompanyGoods>(checkCount);

                //获取已选择的公司-物料
                string comGoods = this.checkedGoodsListBox.CheckedItems[i].ToString();

                //公司抬头名称
                string companyName = comGoods.Split('-')[0];
                //国别&品种
                string goodName = comGoods.Split('-')[1];

                int checkSheetCount = this.checkedSheetListBox.CheckedItems.Count;
                List<string> sheetNameList = new List<string>(checkSheetCount);
                for (int iSheet = 0; iSheet < checkSheetCount; iSheet++)
                {
                    string sheetName = this.checkedSheetListBox.CheckedItems[iSheet].ToString();
                    List<CompanyGoods> companyGoods = getCompanyGoods(companyName, goodName, sheetName);
                    if (keyValues.ContainsKey(comGoods))
                    {
                        List<CompanyGoods> companyGoodsAll = keyValues[comGoods];
                        companyGoodsAll.AddRange(companyGoods);
                    }
                    else
                    {
                        keyValues.Add(comGoods, companyGoods);
                    }
                }
            }
            calGoodsUSD(keyValues);
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
                if (cellAValue.Equals(companyName) && cellHValue.Equals(goodName))
                {
                    //读取第R列提单金额(USD)
                    string cellRValue = ((Range)worksheet.Cells[rowsCount, "R"]).Text.ToString();
                    CompanyGoods companyGood = new CompanyGoods();
                    companyGood.CompanyName = companyName;
                    companyGood.GoodsName = goodName;
                    companyGood.Usd = Double.Parse(cellRValue);
                    companyGood.IsCleared = isCleared;
                    companyGoods.Add(companyGood);
                }
            }
            return companyGoods;
        }

        private void calGoodsUSD(Dictionary<String, List<CompanyGoods>> keyValues) 
        {
            List<StatisticalData> statisticalDatas = new List<StatisticalData>(keyValues.Count);
            //已结关额度
            double clearedQuotaAll = 0.0d;
            //清关中额度
            double customsClearanceAll = 0.0d;
            //总额度
            double customsTotalAll = 0.0d;
            //增量补贴
            double incrementalSubsidyAll = 0.0d;

            foreach (KeyValuePair<String, List<CompanyGoods>> entry in keyValues )
            {
                string key = entry.Key;
                List<CompanyGoods> companyGoodsList = entry.Value;
                if (null != companyGoodsList && companyGoodsList.Any()) 
                {
                    StatisticalData statisticalData = new StatisticalData();
                    statisticalData.CompanyName = key.Split('-')[0];
                    statisticalData.GoodsName = key.Split('-')[1];
                    double clearedQuota = 0d;//已结关额度
                    double customsClearance = 0d;//清关中额度
                    foreach (CompanyGoods cGoods in companyGoodsList) 
                    {
                        if (cGoods.IsCleared) 
                        {
                            clearedQuota += cGoods.Usd;
                        } 
                        else 
                        {
                            customsClearance+= cGoods.Usd;
                        }
                    }
                    statisticalData.ClearedQuota = clearedQuota;//已结关额度
                    statisticalData.CustomsClearance = customsClearance;//清关中额度
                    statisticalData.CustomsTotal = clearedQuota - customsClearance;//总额度
                    Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(statisticalData.CompanyName) && item.GoodsName.Equals(statisticalData.GoodsName));//获取补贴比率对象
                    statisticalData.IncrementalSubsidy = statisticalData.CustomsTotal * ratio.RatioValue;//增量补贴
                    statisticalData.Ratio = ratio.RatioValue;//补贴率
                    statisticalDatas.Add(statisticalData);

                    clearedQuotaAll += clearedQuota;//总已结关额度
                    customsClearanceAll += customsClearance;//总清关中额度
                    customsTotalAll += statisticalData.CustomsTotal; //总额度
                    incrementalSubsidyAll += statisticalData.IncrementalSubsidy;//总增量补贴
                }
            }


        }

        private void checkedGoodsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {           
            CheckedListBox checkedListBox = (CheckedListBox)sender;
            int index = checkedListBox.SelectedIndex;//idx就是当前的选中项序号
            string value = checkedListBox.GetItemText(checkedListBox.Items[index]);
            if (checkedListBox.GetItemChecked(index))
            {
                Ratio ratio = new Ratio();
                ratio.CompanyName = value.Split('-')[0];
                ratio.GoodsName = value.Split('-')[1];
                //被勾选
                ratioList.Add(ratio);//存放补贴比率
                int dataIndex = this.dataGridView.Rows.Add();
                this.dataGridView.Rows[dataIndex].Cells[0].Value = ratio.CompanyName;//公司抬头
                this.dataGridView.Rows[dataIndex].Cells[1].Value = ratio.GoodsName;//品种
                if (dataSet.Contains(value))
                {
                    return;
                }
                else 
                {
                    dataSet.Add(value);//存放补贴比率在dataGridView中第几行
                }
            }
            else 
            {
                //不被勾选
                string companyName = value.Split('-')[0];
                string goodsName = value.Split('-')[1];
                //被勾选
                ratioList.RemoveAll(item=> item.CompanyName.Equals(companyName) && item.GoodsName.Equals(goodsName));//移除补贴比率对象
                foreach (DataGridViewRow rowItem in this.dataGridView.Rows)
                { 
                    string dataCompany = rowItem.Cells[0].Value.ToString();
                    string dataGood = rowItem.Cells[1].Value.ToString();
                    if (companyName.Equals(dataCompany) && goodsName.Equals(dataGood))
                    {
                        this.dataGridView.Rows.Remove(rowItem);
                    }
                }
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //编辑补贴率
            DataGridViewRow rowItem = this.dataGridView.Rows[e.RowIndex];
            string ratioValue = rowItem.Cells[2].Value.ToString();
            if (string.IsNullOrEmpty(ratioValue))
            {
                string dataCompany = rowItem.Cells[0].Value.ToString();
                string dataGood = rowItem.Cells[1].Value.ToString();
                Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(dataCompany) && item.GoodsName.Equals(dataGood));//获取补贴比率对象
                ratio.RatioValue = Double.Parse(ratioValue);
            }
            else 
            {
                string dataCompany = rowItem.Cells[0].Value.ToString();
                string dataGood = rowItem.Cells[1].Value.ToString();
                Ratio ratio = ratioList.Find(item => item.CompanyName.Equals(dataCompany) && item.GoodsName.Equals(dataGood));//获取补贴比率对象
                ratio.RatioValue = 1.0d;
            }
        }

        private void writeExcel(List<StatisticalData> statisticalDatas) 
        {
            //create new xls file
            string file = "C:\\newdoc.xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("First Sheet");
            worksheet.Cells[0, 1] = new Cell((short)1);
            worksheet.Cells[2, 0] = new Cell(9999999);
            worksheet.Cells[3, 3] = new Cell((decimal)3.45);
            worksheet.Cells[2, 2] = new Cell("Text string");
            worksheet.Cells[2, 4] = new Cell("Second string");
            worksheet.Cells[4, 0] = new Cell(32764.5, "#,##0.00");
            worksheet.Cells[5, 1] = new Cell(DateTime.Now, @"YYYY\-MM\-DD");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            workbook.Worksheets.Add(worksheet);
            workbook.Save(file);
            // open xls file
            Workbook book = Workbook.Load(file);
            Worksheet sheet = book.Worksheets[0];
            // traverse cells
            foreach (Pair<Pair<int, int>, Cell> cell in sheet.Cells)
            {
                dgvCells[cell.Left.Right, cell.Left.Left].Value = cell.Right.Value;
            }
            // traverse rows by Index
            for (int rowIndex = sheet.Cells.FirstRowIndex;
      rowIndex <= sheet.Cells.LastRowIndex; rowIndex++)
            {
                Row row = sheet.Cells.GetRow(rowIndex);
                for (int colIndex = row.FirstColIndex;
      colIndex <= row.LastColIndex; colIndex++)
                {
                    Cell cell = row.GetCell(colIndex);
                }
            }
        }
    }
}
