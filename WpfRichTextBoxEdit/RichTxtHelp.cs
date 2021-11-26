using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WpfRichTextBoxEdit
{
    

    public partial class RichTxtHelp
    {
        #region 属性
        struct CellInfo
        {
            /// <summary>
            /// 行
            /// </summary>
            public int R;
            /// <summary>
            /// 列
            /// </summary>
            public int C;
            public double X;
            public double Y;
            public Table Tab;
            /// <summary>
            /// 拖拽方向
            /// </summary>
            public string DragType;
        }
        class CellMark
        {
            public int? Cell { get; }

            public int Row { get; }

            public int? RowSpan { get; }

            public int Col { get; }

            public int? ColSpan { get; }

            public CellMark(int row, int col)
                : this(row, col, null, null, null)
            {

            }

            public CellMark(int row, int col,
                int? rowSpan, int? colSpan,
                int? cell)
            {
                Row = row;
                Col = col;
                RowSpan = rowSpan;
                ColSpan = colSpan;
                Cell = cell;
            }


        }
        private class TreeHelp
        {
            /// <summary>
            /// 在逻辑树中查找指定的元素 <see cref="T"/>
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="relate"></param>
            /// <param name="downward">
            /// 搜索方向，默认值 true（向下搜索）
            /// </param>
            /// <returns>
            /// 返回第一个查到的 <see cref="T"/>。如果，未查到。则返回 <see cref="T"/> 的默认值
            /// </returns>
            public T FirstOrDefault<T>(DependencyObject relate, bool downward = true)
            {
                try
                {
                    if (relate is T r)
                        return r;
                    return downward ? FirstOrDefaultDownward<T>(relate) : FirstOrDefaultUpward<T>(relate);
                }
                catch (Exception ex)
                {
                    throw new Exception("在逻辑树中查找指定的元素失败", ex);
                }
            }

            private T FirstOrDefaultDownward<T>(DependencyObject relate)
            {
                foreach (var child in LogicalTreeHelper.GetChildren(relate))
                {
                    if (child is T result1)
                    {
                        return result1;
                    }

                    if (child is DependencyObject dependencyObject)
                    {
                        if (FirstOrDefaultDownward<T>(dependencyObject) is T result2)
                        {
                            return result2;
                        }
                    }
                }
                return default;
            }

            private T FirstOrDefaultUpward<T>(DependencyObject relate)
            {
                var parent = LogicalTreeHelper.GetParent(relate);
                if (null == parent)
                {
                    return default;
                }
                if (parent is T result1)
                {
                    return result1;
                }

                return FirstOrDefaultUpward<T>(parent);
            }
        }
        TableCell cellEnter;
        TableCell cellLeave;
        TableCell curCell;
        Table tabLeave;
        RichTextBox rtbMain;
        Window window;
        double tableWidth;
        double lineHeight;
        double borderThick; 
        #endregion

        public static int[,] GetArray(DependencyObject obj)
        {
            return (int[,])obj.GetValue(ArrayProperty);
        }

        public static void SetArray(DependencyObject obj, int[,] value)
        {
            obj.SetValue(ArrayProperty, value);
        }

        // Using a DependencyProperty as the backing store for Array.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ArrayProperty =
            DependencyProperty.RegisterAttached("Array", typeof(int[,]), typeof(RichTxtHelp), new PropertyMetadata(null));

        public RichTxtHelp(RichTextBox r, Window w)
        {
            rtbMain = r;
            window = w;
            rtbMain.PreviewKeyDown += RtbMain_KeyDown;
            rtbMain.PreviewMouseMove += RtbMain_PreviewMouseMove;
            rtbMain.SelectionChanged += RtbMain_SelectionChanged;
            rtbMain.ContextMenu = null;
            lineHeight = 20;
            borderThick = 0.5;
        }

        /// <summary>
        /// 选择文本改变，并且鼠标按下时才执行字体等信息刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RtbMain_SelectionChanged(object sender, RoutedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        public void BindAllTabMouse()
        {
            foreach (var block in rtbMain.Document.Blocks.ToList())
            {
                if (block is Table table)
                {
                    BindTabMouse(table);
                }
            }
        }

        public void BindTabMouse(Table table)
        {
            if (table == null)
                return;
            curCell = null;
            cellLeave = null;
            cellEnter = null;
            tabLeave = null;
            foreach (var rowGroup in table.RowGroups)
            {
                var arr=GetMarks(table);
                RichTxtHelp.SetArray(table, arr);
                foreach (var row in rowGroup.Rows)
                {
                    foreach (var c in row.Cells)
                    {
                        //c.Tag = new CellInfo { C = row.Cells.IndexOf(c), R = rowGroup.Rows.IndexOf(row), Tab = table };
                        c.MouseMove -= C_MouseMove; c.MouseMove += C_MouseMove;
                        c.MouseEnter -= C_MouseEnter; c.MouseEnter += C_MouseEnter;
                        c.MouseLeave -= C_MouseLeave; c.MouseLeave += C_MouseLeave;
                    }
                }
            }
            table.MouseLeave -= Table_MouseLeave;
            table.MouseLeave += Table_MouseLeave;
        }

        private void RtbMain_KeyDown(object sender, KeyEventArgs e)
        {
            var key = e.Key == Key.V || e.Key == Key.Z;
            var ctrl = e.KeyboardDevice.Modifiers == ModifierKeys.Control;
            if (ctrl && key||e.Key==Key.Enter)
            {
                new Thread(() => {
                    Thread.Sleep(1000);
                    rtbMain.Dispatcher.Invoke(new Action(() => {
                        if (e.Key == Key.Z)
                            UpdateColumn();
                        BindAllTabMouse(); }));

                }).Start();
            }
        }

        private void C_MouseLeave(object sender, MouseEventArgs e)
        {
            cellLeave = sender as TableCell;
            CellInfo ci = (CellInfo)cellLeave.Tag;
            var p = e.MouseDevice.GetPosition(rtbMain);
            ci.X = p.X;
            ci.Y = p.Y;
            cellLeave.Tag = ci;
            var w = window as MainWindow;
            w.txtMouseLeave.Text = string.Format("CellLeave row:{0}column:{1}x:{2}y:{3}", ci.R, ci.C, p.X, p.Y);

        }

        private void C_MouseEnter(object sender, MouseEventArgs e)
        {
            cellEnter = sender as TableCell;
            CellInfo ci = (CellInfo)cellEnter.Tag;
            var p = e.MouseDevice.GetPosition(rtbMain);
            ci.X = p.X;
            ci.Y = p.Y;
            var w = window as MainWindow;
            w.txtMouseEnter.Text = string.Format("CellEnter row:{0}col:{1}x:{2}y:{3}", ci.R, ci.C, p.X, p.Y);
        }

        private void C_MouseMove(object sender, MouseEventArgs e)
        {
            var p = e.MouseDevice.GetPosition(rtbMain);
            CellInfo ci;
            if (cellEnter == null || cellLeave == null)
                return;
            if (cellEnter == curCell)
                curCell = null;
            CellInfo cLeave = (CellInfo)cellLeave.Tag;
            if (tabLeave != null && cLeave.Tab == tabLeave)
            {
                curCell = cellLeave;
                CellInfo cinfo = (CellInfo)curCell.Tag;
                if (cLeave.C == tabLeave.Columns.Count - 1)
                {
                    cinfo.DragType = "X";
                    curCell.Tag = cinfo;
                    var move = p.X - cinfo.X;
                    if (Math.Abs(move) < 8)
                    {
                        e.MouseDevice.SetCursor(Cursors.SizeWE);
                    }
                    else
                    {
                        tabLeave = null;
                        curCell = null;
                    }
                }
                else if (cLeave.R == tabLeave.RowGroups[0].Rows.Count - 1)
                {
                    cinfo.DragType = "Y";
                    curCell.Tag = cinfo;
                    var move = p.Y - cinfo.Y;
                    if (Math.Abs(move) < 8)
                    {
                        e.MouseDevice.SetCursor(Cursors.SizeNS);
                    }
                    else
                    {
                        tabLeave = null;
                        curCell = null;
                    }

                }
                else
                {
                    tabLeave = null;
                    curCell = null;
                }
                return;
            }

            CellInfo cEnter = (CellInfo)cellEnter.Tag;
            if (cEnter.C < cLeave.C || cEnter.R < cLeave.R)
                curCell = cellEnter;
            else if (cEnter.C > cLeave.C || cEnter.R > cLeave.R)
                curCell = cellLeave;
            else
            {
                curCell = null;
                return;
            }

            ci = (CellInfo)curCell.Tag;

            if (cEnter.C != cLeave.C)
            {
                ci.DragType = "X";
                curCell.Tag = ci;
                var move = p.X - ci.X;
                if (Math.Abs(move) < 8)
                    e.MouseDevice.SetCursor(Cursors.SizeWE);
                else
                    curCell = null;
            }
            else if (cEnter.R != cLeave.R)
            {
                ci.DragType = "Y";
                curCell.Tag = ci;
                var move = p.Y - ci.Y;
                if (Math.Abs(move) < 8)
                    e.MouseDevice.SetCursor(Cursors.SizeNS);
                else
                    curCell = null;
            }
        }

        private void Table_MouseLeave(object sender, MouseEventArgs e)
        {
            tabLeave = sender as Table;
            var p = e.MouseDevice.GetPosition(rtbMain);
            //tabLeave.Tag = new TabInfo() { X = p.X, Y = p.Y };
            var w = window as MainWindow;
            w.txtTabLeave.Text = $"TabLeave  x:{p.X}y:{p.Y}";
        }

        private void RtbMain_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            tableWidth = rtbMain.ActualWidth - 50;
            if (e.MouseDevice.LeftButton == MouseButtonState.Pressed)
            {
                if (curCell != null)
                {
                    var p = e.MouseDevice.GetPosition(rtbMain);
                    CellInfo ci = (CellInfo)curCell.Tag;
                    if (ci.DragType == "X")
                    {
                        e.MouseDevice.SetCursor(Cursors.SizeWE);
                        var move = p.X - ci.X;
                        var column = ci.Tab.Columns[ci.C];
                        var newWidth = move + column.Width.Value;
                        if (newWidth > 10)
                        {
                            column.Width = new GridLength(move + column.Width.Value, GridUnitType.Pixel);
                            ci.X += move;
                            curCell.Tag = ci;
                        }
                    }
                    else if (ci.DragType == "Y")
                    {
                        e.MouseDevice.SetCursor(Cursors.SizeNS);
                        var move = p.Y - ci.Y;
                        //var newHeight = move + curCell.LineHeight;
                        var padTop = curCell.Padding.Top + move / 2;
                        var padBottom = curCell.Padding.Bottom + move / 2;
                        if (padTop > 0&&padBottom>0)
                        {
                            //curCell.LineHeight = curCell.LineHeight + move;
                            curCell.Padding= new Thickness() {Left=curCell.Padding.Left,Right=curCell.Padding.Right,Top= padTop,Bottom=padBottom } ;
                            ci.Y += move;
                            curCell.Tag = ci;
                            var row = curCell.Parent as TableRow;
                            foreach (var c in row.Cells)
                            {
                                if (c != curCell)
                                    //c.LineHeight = curCell.LineHeight;
                                    c.Padding = curCell.Padding;
                            }

                        }
                    }

                }
            }
            else
            {
                if (curCell != null)
                {
                    var p = e.MouseDevice.GetPosition(rtbMain);
                    CellInfo ci = (CellInfo)curCell.Tag;
                    if (Math.Abs(p.X = ci.X) > 8 || Math.Abs(p.Y - ci.Y) > 8)
                        curCell = null;
                }
                    
            }
        }

        public void MergeCell(TextSelection selection)
        {
            var manager = new TreeHelp();
            var startCell = manager.FirstOrDefault<TableCell>(selection.Start.Parent, false);
            var endCell = GetEndCell(selection);
            if (startCell == null || endCell == null)
                return;
            if (startCell == endCell)
                return;
            var startCi = (CellInfo)startCell.Tag;
            var endCi = (CellInfo)endCell.Tag;
            var tab = startCi.Tab;

            var xStart = startCi.C;
            var xEnd = endCi.C + endCell.ColumnSpan - 1;
            var yStart = startCi.R;
            var yEnd = endCi.R + endCell.RowSpan - 1;

            if (!CanMerge(tab, xStart, xEnd, yStart, yEnd))//是否能合并
                return;

            rtbMain.BeginChange();
            #region 执行选中单元格的横向合并
            var colSpan = xEnd - xStart + 1;
            for (int i = yStart; i <= yEnd; i++)
            {
                var row = tab.RowGroups[0].Rows[i];
                TableCell rowStartCell = null;
                for (int j = row.Cells.Count - 1; j >= 0; j--)
                {
                    var c = row.Cells[j];
                    var ci = (CellInfo)c.Tag;
                    if (ci.C > xStart && ci.C <= xEnd)
                    {
                        row.Cells.Remove(c);
                    }
                    else if (ci.C == xStart)
                    {
                        rowStartCell = c;
                    }
                }
                if (rowStartCell != null && colSpan > 0)
                    rowStartCell.ColumnSpan = colSpan;
            }
            #endregion

            # region 执行选中单元格的纵向合并
            var rowSpan = yEnd - yStart + 1;
            for (int i = yStart + 1; i <= yEnd; i++)
            {
                var row = tab.RowGroups[0].Rows[i];
                for (int j = row.Cells.Count - 1; j >= 0; j--)
                {
                    var c = row.Cells[j];
                    var ci = (CellInfo)c.Tag;
                    if (ci.C == xStart)
                    {
                        row.Cells.Remove(c);
                    }

                }
            }
            if (rowSpan > 0)
                startCell.RowSpan = rowSpan;
            #endregion
            rtbMain.EndChange();
            BindTabMouse(tab);
        }

        private bool CanMerge(Table tab, int xStart, int xEnd, int yStart, int yEnd)
        {
            bool r = false;
            var rows = tab.RowGroups[0].Rows;
            int[] colSpans = new int[rows.Count];
            for (int i = yStart; i <= yEnd; i++)
            {
                var row = rows[i];
                for (int j = xStart; j <= xEnd; j++)
                {
                    TableCell cell = null;
                    foreach (var c in row.Cells)
                    {
                        var ci = (CellInfo)c.Tag;
                        if (ci.C == j)
                        {
                            cell = c;
                        }
                    }
                    if (cell != null)
                    {
                        colSpans[i] += cell.ColumnSpan;
                        for (int k = 2; k <= cell.RowSpan; k++)
                        {
                            colSpans[i + k - 1] += cell.ColumnSpan;
                        }
                    }
                }
            }
            r = true;
            for (int i = yStart + 1; i <= yEnd; i++)
            {
                if (colSpans[i] != colSpans[i - 1])
                {
                    r = false;
                    break;
                }

            }
            return r;
        }

        public void SplitCell(TableCell cell)
        {
            if (cell.ColumnSpan<=1&&cell.RowSpan<=1)
                return;
            if (cell == null)
                return;
            var ci = (CellInfo)cell.Tag;
            var xStart = ci.C;
            var xEnd = ci.C + cell.ColumnSpan - 1;
            var yStart = ci.R;
            var yEnd = ci.R + cell.RowSpan - 1;
            rtbMain.BeginChange();
            cell.RowSpan = 1;
            cell.ColumnSpan = 1;
            for (int i = yStart; i <= yEnd; i++)
            {
                var row = ci.Tab.RowGroups[0].Rows[i];
                var index = -1;
                foreach (var c in row.Cells)
                {
                    var cInfo = (CellInfo)c.Tag;
                    if (cInfo.C >= xEnd)
                        break;
                    index = row.Cells.IndexOf(c);
                }
                var column = xStart;
                var count = xEnd - xStart + 1;
                if (i == yStart)
                {
                    count -= 1;//首行，第一格不需要插入
                    column += 1;//首行，从第二格开始插入
                }
                for (int j = 0; j < count; j++)
                {
                    var newCell=CreateCell(Colors.Black);
                    //newCell.Tag = new CellInfo() { C = column, R = i };
                    if (index > row.Cells.Count)
                        row.Cells.Add(newCell);
                    else
                        row.Cells.Insert(index, newCell);
                    column++;
                }
            }
            rtbMain.EndChange();
            BindTabMouse(ci.Tab);
        }

        private bool CanSplit(TextSelection selection)
        {
            bool r = false;
            var manager = new TreeHelp();
            var cell = manager.FirstOrDefault<TableCell>(selection.Start.Parent, false);
            var endCell = GetEndCell(selection);
            if (cell != null && cell == endCell)
            {
                if (cell.RowSpan > 1 || cell.ColumnSpan > 1)
                    r = true;
            }
            return r;
        }

        public int InsertRow(TableCell cell)
        {
            var rowInsert = -1;
            if (cell == null)
                return rowInsert;
            var ci = (CellInfo)cell.Tag;
            var tab = ci.Tab;
            if (tab == null)
                return rowInsert;
            var xStart = 0;
            var xEnd = tab.Columns.Count-1;
            var yStart = ci.R;
            var yEnd = tab.RowGroups[0].Rows.Count-1;
            var array = RichTxtHelp.GetArray(tab);
            rowInsert = yStart;
            for (int i = yStart; i <= yEnd; i++)
            {
                bool canInsert=true;
                for (int j = xStart; j <= xEnd; j++)
                {
                    if (array[j, i] == -1)
                    {
                        canInsert = false;
                        break;//只要有跨行的，就不能插入
                    }
                }
                if (canInsert)
                {
                    break;
                }
                rowInsert++;
            }
            TableRow row = CreateRow(tab.Columns.Count);
            rtbMain.BeginChange();
            if (rowInsert> yEnd)
                tab.RowGroups[0].Rows.Add(row);
            else
                tab.RowGroups[0].Rows.Insert(rowInsert, row);
            rtbMain.EndChange();
            BindTabMouse(tab);
            return rowInsert;
        }

        public int RemoveRow(TableCell cell)
        {
            var rowDel = -1;
            if (cell == null)
                return rowDel;
            var ci = (CellInfo)cell.Tag;
            var tab = ci.Tab;
            if (tab == null)
                return rowDel;
            var xStart = 0;
            var xEnd = tab.Columns.Count - 1;
            var yStart = ci.R;
            var yEnd = tab.RowGroups[0].Rows.Count - 1;
            var array = RichTxtHelp.GetArray(tab);
            rowDel = yStart;
            for (int i = yStart; i <= yEnd; i++)
            {
                bool canRemove = true;
                for (int j = xStart; j <= xEnd; j++)
                {
                    if (array[j, i] == -1)
                    {
                        canRemove = false;
                        break;//只要有跨行的，就不能删除
                    }
                }
                if (canRemove)
                {
                    break;
                }
                rowDel++;
            }
            if (rowDel == yStart)
            {
                rtbMain.BeginChange();
                tab.RowGroups[0].Rows.RemoveAt(rowDel);
                rtbMain.EndChange();
                BindTabMouse(tab);
            }
            else
                rowDel = -1;    
            return rowDel;
        }

        public int InsertCol(TableCell cell)
        {
            var colIndex = -1;
            if (cell == null)
                return colIndex;
            var ci = (CellInfo)cell.Tag;
            var tab = ci.Tab;
            if (tab == null)
                return colIndex;         
            var xStart = ci.C;
            var xEnd = tab.Columns.Count - 1;
            var yStart = 0;
            var yEnd = tab.RowGroups[0].Rows.Count - 1;
            var array = RichTxtHelp.GetArray(tab);
            colIndex = xStart;
            for (int i = xStart; i <= xEnd; i++)
            {
                bool canInsert = true;
                for (int j = yStart; j < yEnd; j++)
                {
                    if (array[i, j] == -1)
                    {
                        canInsert = false;
                        break;
                    }
                }
                if (canInsert)
                {
                    break;
                }
                colIndex++;
            }
            double oldWidth = 0;
            foreach (var col in tab.Columns)
            {
                oldWidth += col.Width.Value;
            }
            rtbMain.BeginChange();
            if (colIndex > xEnd)
            {
                tab.Columns.Add(new TableColumn() { Width = new GridLength(80, GridUnitType.Pixel) });
                foreach (var row in tab.RowGroups[0].Rows)
                {
                    TableCell c = CreateCell(Colors.Black);
                    row.Cells.Add(c);
                }

            }
            else
            {
                tab.Columns.Insert(colIndex, new TableColumn() { Width = new GridLength(80, GridUnitType.Pixel) });
                foreach (var row in tab.RowGroups[0].Rows)
                {
                    var rowIndex = tab.RowGroups[0].Rows.IndexOf(row);
                    var cellIndex = array[colIndex, rowIndex];
                    TableCell c = CreateCell(Colors.Black);
                    row.Cells.Insert(cellIndex, c);
                }
            }
            double newWidth = 0;
            foreach(var col in tab.Columns)
            {
                newWidth += col.Width.Value;
            }
            double rate = 1;
            if (newWidth > tableWidth)
            {
                rate = oldWidth / newWidth;
            }
            if (rate < 1)
            {
                foreach (var col in tab.Columns)
                {
                    col.Width = new GridLength(col.Width.Value * rate, GridUnitType.Pixel);
                }
            }
            rtbMain.EndChange();
            BindTabMouse(tab);
            return colIndex;
        }

        public int RemoveCol(TableCell cell)
        {
            var colIndex = -1;
            if (cell == null)
                return colIndex;
            var ci = (CellInfo)cell.Tag;
            var tab = ci.Tab;
            if (tab == null)
                return colIndex;
            var xStart = ci.C;
            var xEnd = tab.Columns.Count - 1;
            var yStart = 0;
            var yEnd = tab.RowGroups[0].Rows.Count - 1;
            var array = RichTxtHelp.GetArray(tab);
            colIndex = xStart;
            for (int i = xStart; i <= xEnd; i++)
            {
                bool canRemove = true;
                for (int j = yStart; j < yEnd; j++)
                {
                    if (array[i, j] == -1)
                    {
                        canRemove = false;
                        break;
                    }
                }
                if (canRemove)
                {
                    break;
                }
                colIndex++;
            }
            if (colIndex == xStart)
            {
                rtbMain.BeginChange();
                foreach (var row in tab.RowGroups[0].Rows)
                {
                    var rowIndex = tab.RowGroups[0].Rows.IndexOf(row);
                    var cellIndex = array[colIndex, rowIndex];
                    row.Cells.RemoveAt(cellIndex);
                }
                tab.Columns.RemoveAt(colIndex);
                rtbMain.EndChange();
                BindTabMouse(tab);
            }
            else
            {
                colIndex = -1;
            }
            return colIndex;
        }

        private TableCell GetEndCell(TextSelection selection)
        {
            TableCell endCell = null;
            var manager = new TreeHelp();
            TableRow cRow;
            int index;
            var startCell = manager.FirstOrDefault<TableCell>(selection.Start.Parent, false);
            switch (selection.End.Parent)
            {
                case TableRow row:
                    endCell = row.Cells[row.Cells.Count - 1];
                    break;
                case Run run:
                    endCell = manager.FirstOrDefault<TableCell>(run, false);
                    if (endCell == null || startCell == endCell)
                        break;
                    cRow = endCell.Parent as TableRow;
                    index = cRow.Cells.IndexOf(endCell);
                    endCell = cRow.Cells[index - 1];//选中了多格，endcell需要回移才是正确的
                    break;
                case Paragraph paragraph:
                    endCell = manager.FirstOrDefault<TableCell>(paragraph, false);
                    if (endCell == null || startCell == endCell)
                        break;
                    cRow = endCell.Parent as TableRow;
                    index = cRow.Cells.IndexOf(endCell);
                    endCell = cRow.Cells[index - 1];//选中了多格，endcell需要回移才是正确的
                    break;
                case TableCell cell:
                    endCell = cell;
                    if (endCell == null || startCell == endCell)
                        break;
                    cRow = endCell.Parent as TableRow;
                    index = cRow.Cells.IndexOf(endCell);
                    endCell = cRow.Cells[index - 1];//选中了多格，endcell需要回移才是正确的
                    break;
                default:
                    break;
            }
            return endCell;
        }

        public Table InsertTable(int rows, int columns)
        {
            double width = 0;
            Table table = null;
            var paragraph = rtbMain.CaretPosition.Paragraph;
            Paragraph ph1 = new Paragraph();
            ph1.Inlines.Add("");
            if (paragraph!=null)
            {
                var parentTable= new TreeHelp().FirstOrDefault<Table>(paragraph, false);
                if (parentTable != null)
                    return table;
                table = CreateTable(rows, columns, tableWidth);
                paragraph.SiblingBlocks.InsertAfter(paragraph, table);
                rtbMain.Document.Blocks.Add(ph1);
            }
            else
            {
                switch (rtbMain.CaretPosition.Parent)
                {
                    case TableCell cell:
                        break;
                    case ListItem item:
                        break;
                    case FlowDocument doc:
                    case TableRow row:
                        table = CreateTable(rows, columns, tableWidth);
                        rtbMain.Document.Blocks.Add(table);
                        rtbMain.Document.Blocks.Add(ph1);
                        break;
                }
            }
            
            BindTabMouse(table);
            return table;
        }

        public void RemoveTable(TableCell cell)
        {
            if (cell == null)
                return;
            var ci = (CellInfo)cell.Tag;
            if(ci.Tab!=null)
                rtbMain.Document.Blocks.Remove(ci.Tab);
        }

        public void AlignTable(TableCell cell,HorizontalAlignment align)
        {
            if (cell == null)
                return;
            var ci = (CellInfo)cell.Tag;
            if (ci.Tab != null)
            {
                double width=0;
                foreach (var col in ci.Tab.Columns)
                {
                    width += col.Width.Value;
                }
                if (width < tableWidth)
                {
                    if (align == HorizontalAlignment.Left)
                    {
                        ci.Tab.Margin = new Thickness() { Right = tableWidth - width };
                    }
                    else if (align == HorizontalAlignment.Center)
                    {
                        var half = (tableWidth - width) / 2;
                        ci.Tab.Margin = new Thickness() { Left = half,Right=half };
                    }
                    else if (align == HorizontalAlignment.Right)
                    {
                        ci.Tab.Margin = new Thickness() { Left = tableWidth - width };
                    }
                }
            }
        }
    }


    public partial class RichTxtHelp
    {
        public int[,] GetMarks(Table table)
        {
            var marks = new List<CellMark>();
            TableRowGroup rowGroup = table.RowGroups[0];
            var columnCount = GetColumnCount(rowGroup);
            for (int rowIndex = 0; rowIndex < rowGroup.Rows.Count; rowIndex++)
            {
                var colIndex = 0;
                var row = rowGroup.Rows[rowIndex];
                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                {
                    var offset = marks.Where(p => p.Row == rowIndex).ToList();
                    for (int i = 0; i < columnCount; i++)
                    {
                        if (offset.Any(p => p.Col == i))
                        {
                            continue;
                        }

                        colIndex = i;
                        break;
                    }

                    var cell = row.Cells[cellIndex];
                    cell.BorderThickness = new Thickness(borderThick);
                    cell.Tag = new CellInfo() { R = rowIndex, C = colIndex, Tab = table};
                    CreateContextMenu(cell);
                    var mark = new CellMark(rowIndex, colIndex, cell.RowSpan, cell.ColumnSpan, cellIndex);
                    marks.Add(mark);
                    marks.AddRange(GetMarksByColumnSpan(cell, rowIndex, colIndex));
                    marks.AddRange(GetMarksByRowSpan(cell, rowIndex, colIndex));
                }
            }
            int[,] array = new int[columnCount, rowGroup.Rows.Count];
            for (int i = 0; i < columnCount; i++)
            {
                for (int j = 0; j < rowGroup.Rows.Count; j++)
                {
                    foreach (var m in marks)
                    {
                        if (m.Col == i && m.Row == j)
                            array[i, j] =m.Cell==null?-1:(int)m.Cell;
                    }
                }
            }
            return array;
        }

        public void UpdateColumn()
        {
            foreach (var block in rtbMain.Document.Blocks.ToList())
            {
                if (block is Table table)
                {
                    TableRowGroup rowGroup = table.RowGroups[0];
                    var columns = GetColumnCount(rowGroup);
                    if (columns != table.Columns.Count)
                    {
                        table.Columns.Clear();
                        var cellWidth = tableWidth / columns;
                        for (int i = 0; i < columns; i++)
                        {
                            table.Columns.Add(new TableColumn() { Width = new GridLength(cellWidth, GridUnitType.Pixel) });
                        }
                    }
                }
            }
            
        }

        private List<CellMark> GetMarksByRowSpan(TableCell cell, int rowIndex, int columnIndex)
        {
            var marks = new List<CellMark>();
            for (int i = 1; i < cell.RowSpan; i++)
            {
                var cellMark = new CellMark(rowIndex + i, columnIndex);
                marks.Add(cellMark);
                marks.AddRange(GetMarksByColumnSpan(cell, cellMark.Row, columnIndex));
            }

            return marks;
        }

        private List<CellMark> GetMarksByColumnSpan(TableCell cell, int rowIndex, int columnIndex)
        {
            var marks = new List<CellMark>();
            for (int i = 1; i < cell.ColumnSpan; i++)
            {
                var mark = new CellMark(rowIndex, columnIndex + i);
                marks.Add(mark);
            }

            return marks;
        }

        private void SelectionCells(RichTextBox richTextBox, IList<TableCell> cellls)
        {
            var position1 = cellls[0].ContentStart;
            var position2 = cellls[cellls.Count - 1].ContentEnd;
            richTextBox.Selection.Select(position1, position2);
        }

        private int GetColumnCount(TableRowGroup rowGroup)
        {
            var columnCount = 0;
            var row = rowGroup.Rows[0];
            foreach (var cell in row.Cells)
            {
                if (cell.ColumnSpan > 1)
                {
                    columnCount += cell.ColumnSpan;
                }
                else
                {
                    columnCount++;
                }
            }
            return columnCount;
        }

        private Table CreateTable(int rows, int columns, double width)
        {
            var table = new Table();
            table.CellSpacing = 0;
            // 行
            var cellWidth =width>0? width / columns:80;
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add(new TableColumn() { Width = new GridLength(cellWidth, GridUnitType.Pixel) });
            }
            var rowGroup = new TableRowGroup();
            for (int i = 0; i < rows; i++)
            {
                var row = CreateRow(columns);
                rowGroup.Rows.Add(row);
            }
            table.RowGroups.Add(rowGroup);
            return table;
        }

        private TableRow CreateRow(int columns)
        {
            var row = new TableRow();
            for (int i = 0; i < columns; i++)
            {
                var cell = CreateCell(Colors.Black);
                row.Cells.Add(cell);
            }
            return row;
        }

        private TableCell CreateCell(Color color)
        {
            var thick = new Thickness(borderThick);
            var border = new SolidColorBrush() { Color = color };
            var cell = new TableCell() { BorderThickness = thick, BorderBrush = border,LineHeight=lineHeight,AllowDrop=false};
            var paragraph = new Paragraph();
            paragraph.Inlines.Add("");
            cell.Blocks.Add(paragraph);
            return cell;
        }

        private void CreateContextMenu(TableCell c)
        {
            if (c.ContextMenu != null)
            {
                foreach (MenuItem i in c.ContextMenu.Items)
                {
                    var cell = i.CommandParameter as TableCell;
                    if (cell != c)
                        i.CommandParameter = c;
                }
                return;
            }

            ContextMenu menu = new ContextMenu();

            MenuItem item = new MenuItem() {Header= "添加行", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                InsertRow(cell);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "删除行", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                RemoveRow(cell);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "添加列", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                InsertCol(cell);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "删除列", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                RemoveCol(cell);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "合并单元格", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                MergeCell(rtbMain.Selection);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "拆分单元格", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                SplitCell(cell);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "表格靠左", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                AlignTable(cell, HorizontalAlignment.Left);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "表格居中", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                AlignTable(cell, HorizontalAlignment.Center);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "表格靠右", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                AlignTable(cell, HorizontalAlignment.Right);
            };
            menu.Items.Add(item);
            item = new MenuItem() { Header = "删除表格", CommandParameter = c };
            item.Click += (s, e) => {
                var m = s as MenuItem;
                var cell = m.CommandParameter as TableCell;
                RemoveTable(cell);
            };
            menu.Items.Add(item);
            c.ContextMenu = menu;
            c.ContextMenuOpening += C_ContextMenuOpening;
            
        }

        private void C_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            e.Handled = true;
            var c = sender as TableCell;
            if (c.ContextMenu != null)
            {
                c.ContextMenu.IsOpen = true;
            }
        }

        public void InsertImage(BitmapImage image)
        {
            Image i = new Image();
            i.Source = image;
            i.StretchDirection = StretchDirection.DownOnly;
            var p = rtbMain.CaretPosition.Paragraph;
            if (p != null)
            {
                p.Inlines.Add(new InlineUIContainer(i));

            }
            else
            {
                switch (rtbMain.CaretPosition.Parent)
                {
                    case TableCell cell:
                        break;
                    case ListItem item:
                        break;
                    case FlowDocument doc:
                        p = new Paragraph();
                        p.Inlines.Add(new InlineUIContainer(i));
                        doc.Blocks.Add(p);
                        break;
                }
            }
            Keyboard.ClearFocus();
            Keyboard.Focus(rtbMain);
        }
    }

}
