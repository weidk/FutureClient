using FutureClient.Models;
using FutureClient.ViewModel;
using FutureMQClient.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FutureClient
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        MainViewModel dc;
        public MainWindow()
        {
            InitializeComponent();
            this.WindowState = WindowState.Maximized;
            dc = new MainViewModel(this);
            this.DataContext = dc;
            //FutureMQClient.Models.G.db.DbFirst.Where("UserInfo").CreateClassFile(@"D:\workspace\交易接口\总线\FutureClient\Models", "FutureClient");
        }

        private void TradeSideRadioButtonChecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dc != null)
                {
                    var greyConverter = new BrushConverter();
                    var grey = (Brush)greyConverter.ConvertFromString("#DCDCDC");
                    this.buyRadio.Background = grey;
                    this.sellRadio.Background = grey;
                    this.buyOpen.Background = grey;
                    this.sellClose.Background = grey;
                    this.sellOpen.Background = grey;
                    this.buyClose.Background = grey;

                    if (this.buyRadio.IsChecked == true)
                    {
                        dc.TraderColor = "#FF4500";
                        dc.Direction = "1";
                        dc.IsAutoOC = true;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString("#FF4500");
                        this.buyRadio.Background = brush;

                    }
                    else if (this.sellRadio.IsChecked == true)
                    {
                        dc.TraderColor = "#2E8B57";
                        dc.Direction = "2";
                        dc.IsAutoOC = true;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString("#2E8B57");
                        this.sellRadio.Background = brush;
                    }
                    else if (this.buyOpen.IsChecked == true)
                    {
                        string tradeColor = "#FF4500";
                        dc.TraderColor = tradeColor;
                        dc.Direction = "1";
                        dc.Entrust_direction = "1";
                        dc.Futures_direction = "1";
                        dc.IsAutoOC = false;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString(tradeColor);
                        this.buyOpen.Background = brush;
                    }
                    else if (this.buyClose.IsChecked == true)
                    {
                        string tradeColor = "#FF4500";
                        dc.TraderColor = tradeColor;
                        dc.Direction = "1";
                        dc.Entrust_direction = "1";
                        dc.Futures_direction = "2";
                        dc.IsAutoOC = false;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString(tradeColor);
                        this.buyClose.Background = brush;
                    }
                    else if (this.sellOpen.IsChecked == true)
                    {
                        string tradeColor = "#2E8B57";
                        dc.TraderColor = tradeColor;
                        dc.Direction = "2";
                        dc.Entrust_direction = "2";
                        dc.Futures_direction = "1";
                        dc.IsAutoOC = false;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString(tradeColor);
                        this.sellOpen.Background = brush;
                    }
                    else if (this.sellClose.IsChecked == true)
                    {
                        string tradeColor = "#2E8B57";
                        dc.TraderColor = tradeColor;
                        dc.Direction = "2";
                        dc.Entrust_direction = "2";
                        dc.Futures_direction = "2";
                        dc.IsAutoOC = false;
                        var converter = new BrushConverter();
                        var brush = (Brush)converter.ConvertFromString(tradeColor);
                        this.sellClose.Background = brush;
                    }
                }

                ChangePrice();
            }
            catch
            {

            }

            
        }


        private void ChangePrice()
        {

            try
            {
                if (dc != null && dc.SelectedCodeItem != null && dc.ShowHQ != null)
                {
                    if (dc.Direction == "1")
                    {
                        inputPrice.Text = dc.ShowHQ.SelPrice1.ToString();
                    }
                    else
                    {
                        inputPrice.Text = dc.ShowHQ.BuyPrice1.ToString();
                    }
                }
            }
            catch
            {

            }
            
            

        }

        private void codeCoboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            try
            {
                if (dc != null)
                {
                    if (dc.SelectedCodeItem != null)
                    {
                        if (dc.HqDict.ContainsKey(dc.SelectedCodeItem.Code))
                        {
                            dc.ShowHQ = dc.HqDict[dc.SelectedCodeItem.Code];
                        }
                        else
                        {
                            FutureHQ fh = new FutureHQ();
                            fh.SCode = dc.SelectedCodeItem.Code + ".ZJ";
                            dc.ShowHQ = fh;
                        }

                        ChangePrice();
                    }
                }
            }
            catch
            {

            }

            
        }

        private void CollectionViewSource_FilterAccount(object sender, FilterEventArgs e)
        {
            try
            {
                FuturepositionQry_Output t = e.Item as FuturepositionQry_Output;
                if (t != null)
                // If filter is turned on, filter completed items.
                {
                    if (amountCheck.IsChecked == false && int.Parse(t.out_current_amount) <= 0)
                    {
                        e.Accepted = false;
                    }
                    //else if (longCheck.IsChecked == false && t.out_position_flag == "1")
                    //{
                    //    e.Accepted = false;
                    //}
                    //else if (shortCheck.IsChecked == false && t.out_position_flag == "2")
                    //{
                    //    e.Accepted = false;
                    //}
                    else if (bondName.Text != "" && !t.out_stock_code.Contains(bondName.Text.ToUpper()))
                    {
                        e.Accepted = false;
                    }
                    else
                    {
                        e.Accepted = true;
                    }
                }
            }
            catch
            {

            }

        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (AccountDataGrid != null)
                {
                    CollectionViewSource.GetDefaultView(AccountDataGrid.ItemsSource).Refresh();
                }
            }
            catch
            {

            }

            
        }

        private void CollectionViewSource_FilterFinish(object sender, FilterEventArgs e)
        {
            try
            {
                RequestClass t = e.Item as RequestClass;
                if (t != null)
                // If filter is turned on, filter completed items.
                {
                    if (FinishedLongCheck.IsChecked == false && t.direction == "1")
                    {
                        e.Accepted = false;
                    }
                    else if (FinishedShortCheck.IsChecked == false && t.direction == "2")
                    {
                        e.Accepted = false;
                    }
                    else if (FinishedDeal.IsChecked == false && t.orderState.Contains("成"))
                    {
                        e.Accepted = false;
                    }
                    else if (FinishedWithdraw.IsChecked == false && t.orderState.Contains("撤"))
                    {
                        e.Accepted = false;
                    }
                    else if (FinishedInvalid.IsChecked == false && t.orderState.Contains("废"))
                    {
                        e.Accepted = false;
                    }
                    else if (FinishedPairs.IsChecked == false &&  t.addInfo1 != null)
                    {
                        e.Accepted = false;
                    }
                    //else if (FinishedBondName.Text != "" && t.addInfo1=="")
                    //{
                    //    e.Accepted = false;
                    //}
                    else if (FinishedBondName.Text != "" && !t.code.Contains(FinishedBondName.Text.ToUpper()))
                    {
                        e.Accepted = false;
                    }
                    else
                    {
                        e.Accepted = true;
                    }

                }
            }
            catch (Exception ex)
            {

            }
            
        }

        private void FinishedFilter_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (FinishedDataGrid != null)
                {
                    CollectionViewSource.GetDefaultView(FinishedDataGrid.ItemsSource).Refresh();
                }
            }
            catch
            {

            }
            
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            Environment.Exit(0);
        }

        private void CollectionViewSource_FilterToDeals(object sender, FilterEventArgs e)
        {

            try
            {
                RequestClass t = e.Item as RequestClass;
                if (t != null)
                // If filter is turned on, filter completed items.
                {
                    if (TodealsLongCheck.IsChecked == false && t.direction == "1")
                    {
                        e.Accepted = false;
                    }
                    else if (TodealsShortCheck.IsChecked == false && t.direction == "2")
                    {
                        e.Accepted = false;
                    }
                    else if (TodealsPairs.IsChecked == false && t.addInfo1 != null)
                    {
                        e.Accepted = false;
                    }
                    else if (TodealsBondName.Text != "" && !t.code.Contains(TodealsBondName.Text))
                    {
                        e.Accepted = false;
                    }
                    else if ((MainViewModel.FinishedStateText.Contains(t.orderState) || t.entrust_no == "0" || t.error_info != "" || t.direction == null))
                    {
                        e.Accepted = false;
                    }
                    else
                    {
                        e.Accepted = true;
                    }

                }
            }
            catch
            {

            }

            
        }

        private void ToDealsFilter_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (AliveOrdersDataGrid != null)
                {
                    CollectionViewSource.GetDefaultView(AliveOrdersDataGrid.ItemsSource).Refresh();
                }
            }
            catch
            {

            }
            
        }

        private void MinusChangedEvent(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (dc != null)
                {
                    dc.ChangeMinus(false);
                    dc.ChangeMinus(true);
                }
            }
            catch
            {

            }

            
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dc != null)
                {
                    dc.ChangeMinus(false);
                    dc.ChangeMinus(true);
                }
            }
            catch
            {

            }
            
        }

        private void Pairs_Filter(object sender, FilterEventArgs e)
        {
            try
            {
                PairOrders t = e.Item as PairOrders;
                if (t != null)
                // If filter is turned on, filter completed items.
                {
                    //if (PairIsFinish.IsChecked == false && (t.IsFinish || t.IsDelete))
                    if (PairIsFinish.IsChecked == false && t.IsFinish)
                    {
                        e.Accepted = false;
                    }
                    else if (PairIsUNFinish.IsChecked == false && !t.IsFinish)
                    {
                        e.Accepted = false;
                    }
                    else
                    {
                        e.Accepted = true;
                    }

                }
            }
            catch
            {

            }


            
        }

        private void Pairs_Changed(object sender, RoutedEventArgs e)
        {
            
            try
            {
                if (PairsDataGrid != null)
                {
                    CollectionViewSource.GetDefaultView(PairsDataGrid.ItemsSource).Refresh();
                }
            }
            catch
            {

            }
        }

        public void PairsChange()
        {
            try
            {
                if (PairsDataGrid != null)
                {
                    CollectionViewSource.GetDefaultView(PairsDataGrid.ItemsSource).Refresh();
                }
            }
            catch
            {

            }
            
        }

        private void AccountDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var dg = sender as DataGrid;

                if (dg == null) return;
                var selectedItem = dg.CurrentItem as FuturepositionQry_Output;
                if (selectedItem != null)
                {
                    if (selectedItem.out_position_flag == "1")
                    {
                        sellRadio.IsChecked = true;
                    }
                    else
                    {
                        buyRadio.IsChecked = true;
                    }
                    TradeSideRadioButtonChecked(sender, e);
                    codeCoboBox.Text = selectedItem.out_stock_code;
                    codeCoboBox.SelectedItem = new FutureCode() { Code = selectedItem.out_stock_code };
                    int amt = int.Parse(selectedItem.out_enable_amount);
                    if (amt > 50)
                    {
                        DirectionAmt.Text = "50";
                    }
                    else
                    {
                        DirectionAmt.Text = selectedItem.out_enable_amount;
                    }

                }
            }
            catch
            {

            }
            

        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var dg = sender as DataGrid;

            if (dg == null) return;
            var selectedItem = dg.CurrentItem as FutureHQMvvM;
            if (selectedItem != null)
            {
                if (dg.CurrentColumn.SortMemberPath == "BuyPrice1" || dg.CurrentColumn.SortMemberPath == "BuyVol1")
                {
                    sellRadio.IsChecked = true;
                }
                if (dg.CurrentColumn.SortMemberPath == "SelPrice1" || dg.CurrentColumn.SortMemberPath == "SelVol1")
                {
                    buyRadio.IsChecked = true;
                }
                TradeSideRadioButtonChecked(sender, e);
                char[] charsToTrim = { '.', 'Z', 'J' };
                string code = selectedItem.SCode.TrimEnd(charsToTrim);
                codeCoboBox.Text = code;
                codeCoboBox.SelectedItem = new FutureCode() { Code = code };

            }
            inputPrice.Focus();
        }


        private void inputPrice_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (dc != null && dc.SelectedCodeItem != null && dc.ShowHQ != null)
                {
                    double orginalPrice = double.Parse(inputPrice.Text);
                    if (e.KeyStates == Keyboard.GetKeyStates(Key.Right))
                    {
                        double newPrice = orginalPrice + 0.005d;
                        inputPrice.Text = newPrice.ToString();
                    }
                    if (e.KeyStates == Keyboard.GetKeyStates(Key.Left))
                    {
                        double newPrice = orginalPrice - 0.005d;
                        inputPrice.Text = newPrice.ToString();
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var tc = sender as TabControl; //The sender is a type of TabControl...

                if (tc != null && dc != null)
                {
                    if (PairItem != null && PairItem.IsSelected)
                    {
                        dc.MainColWidth_0 = "0*";
                        dc.MainColWidth_1 = "24*";
                    }
                    else
                    {
                        dc.MainColWidth_0 = "24*";
                        dc.MainColWidth_1 = "0*";
                    }
                }
            }
            catch
            {

            }
            
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<RequestClass> expireList = dc.ExpiredlatestIpList.ToList();
                expireList.ForEach(o => o.direction = o.direction == "1" ? "多" : "空");
                var toExportList = expireList.Select(o => {
                    return new
                    {
                        类型 = o.businessType,
                        代码 = o.code,
                        方向 = o.direction,
                        开平 = o.futures_direction,
                        委托价 = o.entrust_price,
                        成交价 = o.deal_price,
                        委托数量 = o.entrust_amount,
                        成交数量 = o.deal_amount,
                        状态 = o.orderState,
                        配对信息 = o.addInfo1,
                        时间 = o.addInfo3,
                        错误信息 = o.error_info
                    };
                }
                ).ToList();


                var dt = ToDataTable(toExportList);
                String FlName = @"\委托流水_"
                + DateTime.Now.Year.ToString() + "-"
                + DateTime.Now.Month.ToString() + "-"
                + DateTime.Now.Day.ToString() + ".csv";

                TableToCsv(dt, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + FlName);

                //List<RequestClass> expireList = dc.ExpiredlatestIpList.ToList();
                //var dt = ToDataTable(expireList);

                //ToExcel(FinishedDataGrid);

                //bool v = Export(FinishedDataGrid, "委托流水");
                //if (v == true)
                //{
                //    MessageBox.Show("导出文件保存成功！", "成功！");
                //}
                //else MessageBox.Show("导出文件保存失败！", "警告！");
            }
            catch
            {
                MessageBox.Show("导出文件保存失败！", "警告！");
            }
            
        }

        #region EXCEL封装方法
        public bool Export(DataGrid dataGrid, string excelTitle)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < dataGrid.Columns.Count; i++)
            {
                if (dataGrid.Columns[i].Visibility == System.Windows.Visibility.Visible)//只导出可见列  
                {
                    dt.Columns.Add(dataGrid.Columns[i].Header.ToString());//构建表头  
                }
            }

            for (int i = 0; i < dataGrid.Items.Count; i++)
            {
                int columnsIndex = 0;
                DataRow row = dt.NewRow();
                for (int j = 0; j < dataGrid.Columns.Count; j++)
                {
                    if (dataGrid.Columns[j].Visibility == System.Windows.Visibility.Visible)
                    {
                        if (dataGrid.Items[i] != null && (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock) != null)//填充可见列数据  
                        {
                            row[columnsIndex] = (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock).Text.ToString();
                            //row[columnsIndex] = dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]);
                        }
                        else row[columnsIndex] = "";
     
                        columnsIndex++;
                    }
                }
                dt.Rows.Add(row);
            }

            if (ExcelExport(dt, excelTitle) != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string ExcelExport(System.Data.DataTable DT, string title)
        {
            try
            {
                //创建Excel
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelBook = ExcelApp.Workbooks.Add(System.Type.Missing);
                //创建工作表（即Excel里的子表sheet） 1表示在子表sheet1里进行数据导出
                Microsoft.Office.Interop.Excel.Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[1];

                //如果数据中存在数字类型 可以让它变文本格式显示
                ExcelSheet.Cells.NumberFormat = "@";

                //设置工作表名
                ExcelSheet.Name = title;

                //设置Sheet标题
                string start = "A1";
                string end = ChangeASC(DT.Columns.Count) + "1";

                Microsoft.Office.Interop.Excel.Range _Range = (Microsoft.Office.Interop.Excel.Range)ExcelSheet.get_Range(start, end);
                _Range.Merge(0);                     //单元格合并动作(要配合上面的get_Range()进行设计)
                _Range = (Microsoft.Office.Interop.Excel.Range)ExcelSheet.get_Range(start, end);
                _Range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                _Range.Font.Size = 22; //设置字体大小
                _Range.Font.Name = "宋体"; //设置字体的种类 
                ExcelSheet.Cells[1, 1] = title;    //Excel单元格赋值
                _Range.EntireColumn.AutoFit(); //自动调整列宽

                //写表头
                for (int m = 1; m <= DT.Columns.Count; m++)
                {
                    ExcelSheet.Cells[2, m] = DT.Columns[m - 1].ColumnName.ToString();

                    start = "A2";
                    end = ChangeASC(DT.Columns.Count) + "2";

                    _Range = (Microsoft.Office.Interop.Excel.Range)ExcelSheet.get_Range(start, end);
                    _Range.Font.Size = 15; //设置字体大小
                    _Range.Font.Bold = true;//加粗
                    _Range.Font.Name = "宋体"; //设置字体的种类  
                    _Range.EntireColumn.AutoFit(); //自动调整列宽 
                    _Range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }

                //写数据
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    for (int j = 1; j <= DT.Columns.Count; j++)
                    {
                        //Excel单元格第一个从索引1开始
                        // if (j == 0) j = 1;
                        ExcelSheet.Cells[i + 3, j] = DT.Rows[i][j - 1].ToString();
                    }
                }

                //表格属性设置
                for (int n = 0; n < DT.Rows.Count + 1; n++)
                {
                    start = "A" + (n + 3).ToString();
                    end = ChangeASC(DT.Columns.Count) + (n + 3).ToString();

                    //获取Excel多个单元格区域
                    _Range = (Microsoft.Office.Interop.Excel.Range)ExcelSheet.get_Range(start, end);

                    _Range.Font.Size = 12; //设置字体大小
                    _Range.Font.Name = "宋体"; //设置字体的种类

                    _Range.EntireColumn.AutoFit(); //自动调整列宽
                    _Range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; //设置字体在单元格内的对其方式 _Range.EntireColumn.AutoFit(); //自动调整列宽 
                }

                ExcelApp.DisplayAlerts = false; //保存Excel的时候，不弹出是否保存的窗口直接进行保存 

                //弹出保存对话框,并保存文件
                Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog();
                sfd.DefaultExt = ".xlsx";
                sfd.Filter = "Office 2007 File|*.xlsx|Office 2000-2003 File|*.xls|所有文件|*.*";
                if (sfd.ShowDialog() == true)
                {
                    if (sfd.FileName != "")
                    {
                        ExcelBook.SaveAs(sfd.FileName);  //将其进行保存到指定的路径
                                                         //    MessageBox.Show("导出文件已存储为: " + sfd.FileName, "温馨提示");
                    }
                }

                //释放可能还没释放的进程
                ExcelBook.Close();
                ExcelApp.Quit();
                // PubHelper.Instance.KillAllExcel(ExcelApp);

                return sfd.FileName;
            }
            catch
            {
                //   MessageBox.Show("导出文件保存失败！", "警告！");
                return null;
            }
        }

        private string ChangeASC(int count)
        {
            string ascstr = "";
            switch (count)
            {
                case 1:
                    ascstr = "A";
                    break;
                case 2:
                    ascstr = "B";
                    break;
                case 3:
                    ascstr = "C";
                    break;
                case 4:
                    ascstr = "D";
                    break;
                case 5:
                    ascstr = "E";
                    break;
                case 6:
                    ascstr = "F";
                    break;
                case 7:
                    ascstr = "G";
                    break;
                case 8:
                    ascstr = "H";
                    break;
                case 9:
                    ascstr = "I";
                    break;
                case 10:
                    ascstr = "J";
                    break;
                case 11:
                    ascstr = "K";
                    break;
                case 12:
                    ascstr = "L";
                    break;
                case 13:
                    ascstr = "M";
                    break;
                case 14:
                    ascstr = "N";
                    break;
                case 15:
                    ascstr = "O";
                    break;
                case 16:
                    ascstr = "P";
                    break;
                case 17:
                    ascstr = "Q";
                    break;
                case 18:
                    ascstr = "R";
                    break;
                case 19:
                    ascstr = "S";
                    break;
                case 20:
                    ascstr = "T";
                    break;
                default:
                    ascstr = "U";
                    break;
            }
            return ascstr;
        }



        #endregion

        #region 导出csv
        public static void TableToCsv(DataGrid dataGrid, string filePath)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < dataGrid.Columns.Count; i++)
            {
                if (dataGrid.Columns[i].Visibility == System.Windows.Visibility.Visible)//只导出可见列  
                {
                    dt.Columns.Add(dataGrid.Columns[i].Header.ToString());//构建表头  
                }
            }

            for (int i = 0; i < dataGrid.Items.Count; i++)
            {
                int columnsIndex = 0;
                DataRow row = dt.NewRow();
                for (int j = 0; j < dataGrid.Columns.Count; j++)
                {
                    if (dataGrid.Columns[j].Visibility == System.Windows.Visibility.Visible)
                    {
                        if (dataGrid.Items[i] != null && (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock) != null)//填充可见列数据  
                        {
                            row[columnsIndex] = (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock).Text.ToString();
                            //row[columnsIndex] = dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]);
                        }
                        else row[columnsIndex] = "";

                        columnsIndex++;
                    }
                }
                dt.Rows.Add(row);
            }
            FileInfo fi = new FileInfo(filePath);
            string path = fi.DirectoryName;
            string name = fi.Name;
            //\/:*?"<>|
            //把文件名和路径分别取出来处理
            name = name.Replace(@"\", "＼");
            name = name.Replace(@"/", "／");
            name = name.Replace(@":", "：");
            name = name.Replace(@"*", "＊");
            name = name.Replace(@"?", "？");
            name = name.Replace(@"<", "＜");
            name = name.Replace(@">", "＞");
            name = name.Replace(@"|", "｜");
            string title = "";

            FileStream fs = new FileStream(path + "\\" + name, FileMode.Create);
            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                title += dt.Columns[i].ColumnName + ",";
            }
            title = title.Substring(0, title.Length - 1) + "\n";
            sw.Write(title);

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                string line = "";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    line += row[i].ToString().Replace(",", "") + ",";
                }
                line = line.Substring(0, line.Length - 1) + "\n";

                sw.Write(line);
            }




            sw.Close();
            fs.Close();
            MessageBox.Show(name + "已保存到桌面");
        }

        public static void TableToCsv(DataTable dt, string filePath)
        {
            FileInfo fi = new FileInfo(filePath);
            string path = fi.DirectoryName;
            string name = fi.Name;
            //\/:*?"<>|
            //把文件名和路径分别取出来处理
            name = name.Replace(@"\", "＼");
            name = name.Replace(@"/", "／");
            name = name.Replace(@":", "：");
            name = name.Replace(@"*", "＊");
            name = name.Replace(@"?", "？");
            name = name.Replace(@"<", "＜");
            name = name.Replace(@">", "＞");
            name = name.Replace(@"|", "｜");
            string title = "";

            FileStream fs = new FileStream(path + "\\" + name, FileMode.Create);
            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                title += dt.Columns[i].ColumnName + ",";
            }
            title = title.Substring(0, title.Length - 1) + "\n";
            sw.Write(title);

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                string line = "";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    line += row[i].ToString().Replace(",", "") + ",";
                }
                line = line.Substring(0, line.Length - 1) + "\n";

                sw.Write(line);
            }




            sw.Close();
            fs.Close();
            MessageBox.Show(name + "已保存到桌面");
        }
        #endregion


        #region 保存到excel表 会根据datatable表 自动生成表头

        public void ToExcel(DataGrid dataGrid)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < dataGrid.Columns.Count; i++)
            {
                if (dataGrid.Columns[i].Visibility == System.Windows.Visibility.Visible)//只导出可见列  
                {
                    dt.Columns.Add(dataGrid.Columns[i].Header.ToString());//构建表头  
                }
            }

            for (int i = 0; i < dataGrid.Items.Count; i++)
            {
                int columnsIndex = 0;
                DataRow row = dt.NewRow();
                for (int j = 0; j < dataGrid.Columns.Count; j++)
                {
                    if (dataGrid.Columns[j].Visibility == System.Windows.Visibility.Visible)
                    {
                        if (dataGrid.Items[i] != null && (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock) != null)//填充可见列数据  
                        {
                            row[columnsIndex] = (dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]) as TextBlock).Text.ToString();
                            //row[columnsIndex] = dataGrid.Columns[j].GetCellContent(dataGrid.Items[i]);
                        }
                        else row[columnsIndex] = "";

                        columnsIndex++;
                    }
                }
                dt.Rows.Add(row);
            }



            //方法来源
            //http://blog.csdn.net/sanjiawan/article/details/6818921

            #region 初始化 生成Worksheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWB = excelApp.Workbooks.Add(System.Type.Missing);    //创建工作簿（WorkBook：即Excel文件主体本身）  
            Microsoft.Office.Interop.Excel.Worksheet excelWS = (Microsoft.Office.Interop.Excel.Worksheet)excelWB.Worksheets[1];   //创建工作表（即Excel里的子表sheet） 1表示在子表sheet1里进行数据导出  
            #endregion

            #region 新增一行用于保存标题
            DataRow dr = dt.NewRow();
            //sdt表中有的数据是int类型的这样当插入标题行的时候会提示类型不同 
            //所以在这里只是插入一个空行 然后标题列在excel表里设置
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    //    MessageBox.Show(dr[i].GetType().ToString());
            //    //dr[i] = 1;

            //}
            dt.Rows.InsertAt(dr, 0);
            #endregion

            #region 把datatable表中数据写入到 Worksheet中
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    excelWS.Cells[i + 1, j + 1] = dt.Rows[i][j].ToString();   //Excel单元格第一个从索引1开始  
                }
            }
            #endregion

            #region 通过excel的方法编辑标题行
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                excelWS.Cells[1, i + 1] = dt.Columns[i].ColumnName; ; //Excel单元格赋值
            }
            #endregion


            #region 打开保存框
            //打开 SaveFileDialog 框
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            //定义导出的文件名称
            String FlName = "委托流水_"
                + DateTime.Now.Year.ToString() + "-"
                + DateTime.Now.Month.ToString() + "-"
                + DateTime.Now.Day.ToString();
            dlg.FileName = FlName;
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension
            #endregion

            #region 保存到文件
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string filename = dlg.FileName;

                if (dlg.FileName == "")
                {
                    MessageBox.Show("请输入保存文件名！");
                }
                else
                {
                    excelWB.SaveAs(filename);  //将其进行保存到指定的路径  
                    excelWB.Close();
                    excelApp.Quit();

                    //释放可能还没释放的进程  
                    //该方法会导致 如果你已经打开一个excel啦 执行该方法时 会导致你打开的这个excel关闭。
                    //KillAllExcel(excelApp);

                    //显示保存的地址
                    MessageBox.Show("您需要的Excel文件已经保存到" + dlg.FileName);
                }
            }
            #endregion
        }


        public void ToExcel(DataTable dt)
        {
            #region 初始化 生成Worksheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWB = excelApp.Workbooks.Add(System.Type.Missing);    //创建工作簿（WorkBook：即Excel文件主体本身）  
            Microsoft.Office.Interop.Excel.Worksheet excelWS = (Microsoft.Office.Interop.Excel.Worksheet)excelWB.Worksheets[1];   //创建工作表（即Excel里的子表sheet） 1表示在子表sheet1里进行数据导出  
            #endregion

            #region 新增一行用于保存标题
            DataRow dr = dt.NewRow();
            //sdt表中有的数据是int类型的这样当插入标题行的时候会提示类型不同 
            //所以在这里只是插入一个空行 然后标题列在excel表里设置
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    //    MessageBox.Show(dr[i].GetType().ToString());
            //    //dr[i] = 1;

            //}
            dt.Rows.InsertAt(dr, 0);
            #endregion

            #region 把datatable表中数据写入到 Worksheet中
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    excelWS.Cells[i + 1, j + 1] = dt.Rows[i][j].ToString();   //Excel单元格第一个从索引1开始  
                }
            }
            #endregion

            #region 通过excel的方法编辑标题行
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                excelWS.Cells[1, i + 1] = dt.Columns[i].ColumnName; ; //Excel单元格赋值
            }
            #endregion


            #region 打开保存框
            //打开 SaveFileDialog 框
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            //定义导出的文件名称
            String FlName = "委托流水_"
                + DateTime.Now.Year.ToString() + "-"
                + DateTime.Now.Month.ToString() + "-"
                + DateTime.Now.Day.ToString();
            dlg.FileName = FlName;
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension
            #endregion

            #region 保存到文件
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string filename = dlg.FileName;

                if (dlg.FileName == "")
                {
                    MessageBox.Show("请输入保存文件名！");
                }
                else
                {
                    excelWB.SaveAs(filename);  //将其进行保存到指定的路径  
                    excelWB.Close();
                    excelApp.Quit();

                    //释放可能还没释放的进程  
                    //该方法会导致 如果你已经打开一个excel啦 执行该方法时 会导致你打开的这个excel关闭。
                    //KillAllExcel(excelApp);

                    //显示保存的地址
                    MessageBox.Show("您需要的Excel文件已经保存到" + dlg.FileName);
                }
            }
            #endregion
        }

        public DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dt = new DataTable();
            for (int i = 0; i < properties.Count; i++)
            {
                PropertyDescriptor property = properties[i];
                dt.Columns.Add(property.Name, property.PropertyType);
            }
            object[] values = new object[properties.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = properties[i].GetValue(item);
                }
                dt.Rows.Add(values);
            }
            return dt;
        }



        #endregion

    }
}
