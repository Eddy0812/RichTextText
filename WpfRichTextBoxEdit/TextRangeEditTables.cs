using System;
using System.Linq;
using System.Reflection;
using System.Windows.Documents;

namespace WpfRichTextBoxEdit
{
    /// <summary>
    /// 表格相关操作
    /// </summary>
    public class TextRangeEditTables
    {
        //-------------------------------------------------------------------------------------------------\\
        //
        //  通过反射获取到表格操作方法，并调用之
        //
        //  详细请查看.NET Framwork WPF RichTextBox 相关源代码
        //  http://referencesource.microsoft.com/#PresentationFramework/src/Framework/System/Windows/Documents/TextRangeEditTables.cs
        //  http://referencesource.microsoft.com/#PresentationFramework/src/Framework/System/Windows/Documents/TextRange.cs
        //
        //-------------------------------------------------------------------------------------------------//

        #region 表格相关操作

        /// <summary>
        /// 获取选中单元格的第一个（左上角）和最后一个（右下角）单元格
        /// </summary>
        /// <param name="selection">RichTextBox.Section</param>
        /// <param name="startCell"></param>
        /// <param name="endCell"></param>
        /// <returns></returns>
        public static bool GetSelectedCells(TextSelection selection, out TableCell startCell, out TableCell endCell)
        {
            startCell = null;
            endCell = null;

            #region 函数原型
            /********************************************************************************************\
            /// <summary>
            /// From two text positions finds out table elements involved
            /// into building potential table range.
            /// </summary>
            /// <param name="anchorPosition">
            /// Position where selection starts. The cell at this position (if any)
            /// must be included into a range unconditionally.
            /// </param>
            /// <param name="movingPosition">
            /// A position opposite to an anchorPosition.
            /// </param>
            /// <param name="includeCellAtMovingPosition">
            /// <see ref="TextRangeEditTables.BuildTableRange"/>
            /// </param>
            /// <param name="anchorCell">
            /// The cell at anchor position. Returns not null only if a range is not crossing table
            /// boundary. Returns null if the range does not cross any TableCell boundary at all
            /// or if cells crossed belong to a table whose boundary is crossed by a range.
            /// In other words, anchorCell and movingCell are either both nulls or both non-nulls.
            /// </param>
            /// <param name="movingCell">
            /// The cell at the movingPosition.  Returns not null only if a range is not crossing table
            /// boundary. Returns null if the range does not cross any TableCell boundary at all
            /// or if cells crossed belong to a table whose boundary is crossed by a range.
            /// In other words, anchorCell and movingCell are either both nulls or both non-nulls.
            /// </param>
            /// <param name="anchorRow"></param>
            /// <param name="movingRow"></param>
            /// <param name="anchorRowGroup"></param>
            /// <param name="movingRowGroup"></param>
            /// <param name="anchorTable"></param>
            /// <param name="movingTable"></param>
            /// <returns>
            /// True if at least one structural unit was found.
            /// False if no structural units were crossed by either startPosition or endPosition
            /// (up to their commin ancestor element).
            /// </returns>
            private static bool IdentifyTableElements(
                TextPointer anchorPosition, TextPointer movingPosition,
                bool includeCellAtMovingPosition,
                out TableCell anchorCell, out TableCell movingCell,
                out TableRow anchorRow, out TableRow movingRow,
                out TableRowGroup anchorRowGroup, out TableRowGroup movingRowGroup,
                out Table anchorTable, out Table movingTable)
             \********************************************************************************************/
            #endregion

            //System.Windows.Documents.TextRangeEditTables
            Type objectType = (from asm in AppDomain.CurrentDomain.GetAssemblies()
                               from type in asm.GetTypes()
                               where type.IsClass
                               && asm.ManifestModule.Name == "PresentationFramework.dll"
                               && type.Name == "TextRangeEditTables"
                               select type).Single();
            //MethodInfo info = objectType.GetMethod("IdentifyTableElements", BindingFlags.NonPublic | BindingFlags.Static);
            MethodInfo info = getNonPublicMethodInfo(objectType, "IdentifyTableElements");
            if (info != null)
            {
                object[] param = new object[11];
                param[0] = selection.Start;
                param[1] = selection.End;
                param[2] = false;

                object result = info.Invoke(null, param);
                startCell = param[3] as TableCell;
                endCell = param[4] as TableCell;
                return (bool)result;
            }
            return false;
        }

        /// <summary>
        /// 选中单元格是否能合并
        /// </summary>
        /// <param name="selection">RichTextBox.Section</param>
        /// <returns></returns>
        public static bool CanMergeCellRange(TextSelection selection)
        {
            TableCell startCell = null;
            TableCell endCell = null;
            Type objectType = (from asm in AppDomain.CurrentDomain.GetAssemblies()
                               from type in asm.GetTypes()
                               where type.IsClass
                               && asm.ManifestModule.Name == "PresentationFramework.dll"
                               && type.Name == "TextRangeEditTables"
                               select type).Single();
            //MethodInfo info = objectType.GetMethod("CanMergeCellRange", BindingFlags.NonPublic | BindingFlags.Static);
            MethodInfo info = getNonPublicMethodInfo(objectType, "CanMergeCellRange");
            if (info != null)
            {
                GetSelectedCells(selection, out startCell, out endCell);
                if (startCell != null && endCell != null)
                {
                    int startColumnIndex = (int)getPrivateProperty<TableCell>(startCell, "ColumnIndex");
                    int endColumnIndex = (int)getPrivateProperty<TableCell>(endCell, "ColumnIndex");
                    int startRowIndex = (int)getPrivateProperty<TableCell>(startCell, "RowIndex");
                    int endRowIndex = (int)getPrivateProperty<TableCell>(endCell, "RowIndex");
                    TableRowGroup rowGroup = getPrivateProperty<TableRow>(startCell.Parent, "RowGroup") as TableRowGroup;
                    return (bool)info.Invoke(null, new object[] {
                        rowGroup,                               // RowGroup
                        startRowIndex,                          // topRow
                        endRowIndex + endCell.RowSpan - 1,      // bottomRow
                        startColumnIndex,                       // leftColumn
                        endColumnIndex + endCell.ColumnSpan - 1 // rightColumn
                     });
                }
            }
            return false;
        }

        /// <summary>
        /// 合并选中表格
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static TextRange MergeCells(TextRange selection)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("MergeCells");
            if (mInfo != null)
            {
                return mInfo.Invoke(selection, null) as TextRange;
            }
            return null;
        }

        /// <summary>
        /// 拆分表格（好像还有问题。。。）
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="splitCountHorizontal"></param>
        /// <param name="splitCountVertical"></param>
        /// <returns></returns>
        public static TextRange SplitCell(TextRange selection, int splitCountHorizontal, int splitCountVertical)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("SplitCell");
            if (mInfo != null)
            {
                return mInfo.Invoke(selection, new object[] { splitCountHorizontal, splitCountVertical }) as TextRange;
            }
            return null;
        }

        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="rowCount">行数</param>
        /// <param name="columnCount">列数</param>
        /// <returns></returns>
        public static TextRange InsertTable(TextRange selection, int rowCount, int columnCount)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("InsertTable");
            if (mInfo != null)
            {
                return mInfo.Invoke(selection, new object[] { rowCount, columnCount }) as TextRange;
            }
            return null;
        }

        /// <summary>
        /// 在光标下插入行
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="rowCount">行数</param>
        /// <returns></returns>
        public static TextRange InsertRows(TextRange selection, int rowCount)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("InsertRows");
            if (mInfo != null)
            {
                return mInfo.Invoke(selection, new object[] { rowCount }) as TextRange;
            }
            return null;
        }

        /// <summary>
        /// 删除选中行
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static bool DeleteRows(TextRange selection)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("DeleteRows");
            if (mInfo != null)
            {
                return (bool)mInfo.Invoke(selection, null);
            }
            return false;
        }

        /// <summary>
        /// 在光标右边插入列
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="columnCount">列数</param>
        /// <returns></returns>
        public static TextRange InsertColumns(TextRange selection, int columnCount)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("InsertColumns");
            if (mInfo != null)
            {
                return mInfo.Invoke(selection, new object[] { columnCount }) as TextRange;
            }
            return null;
        }

        /// <summary>
        /// 删除选中列
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static bool DeleteColumns(TextRange selection)
        {
            MethodInfo mInfo = getNonPublicMethodInfo<TextRange>("DeleteColumns");
            if (mInfo != null)
            {
                return (bool)mInfo.Invoke(selection, null);
            }
            return false;
        }

        /// <summary>
        /// 获取类中私有方法
        /// </summary>
        /// <param name="type"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        private static MethodInfo getNonPublicMethodInfo(Type type, string methodName)
        {
            MethodInfo mInfo = type
                .GetMethod(methodName,
                BindingFlags.NonPublic
                | BindingFlags.Static
                | BindingFlags.Instance);
            return mInfo;
        }

        /// <summary>
        /// 获取类中私有方法
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="methodName"></param>
        /// <returns></returns>
        private static MethodInfo getNonPublicMethodInfo<T>(string methodName)
            where T : class
        {
            return getNonPublicMethodInfo(typeof(T), methodName);
        }

        /// <summary>
        /// 获取私有属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private static object getPrivateProperty<T>(object instance, string propertyName)
            where T : class
        {
            object result = null;
            PropertyInfo pInfo = typeof(T).GetProperty(propertyName, BindingFlags.NonPublic | BindingFlags.Instance);
            if (pInfo != null)
            {
                result = pInfo.GetValue(instance, null);
            }
            return result;
        }

        #endregion
    }
}