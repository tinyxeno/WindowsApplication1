using System.Collections;
using System.Windows.Forms;

namespace WindowsApplication1
{
    public class ListViewColumnSorter : IComparer
    {
        /// <summary>
        /// 並べ替える列
        /// </summary>
        private int m_column;
        /// <summary>
        /// 並べ替える方向
        /// </summary>
        private SortOrder m_direction;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ListViewColumnSorter()
        {
            Clear();
        }

        /// <summary>
        /// 並べ替える列位置
        /// </summary>
        public int SortColumn
        {
            set
            {
                m_column = value;
            }
            get
            {
                return m_column;
            }
        }

        /// <summary>
        /// 並べ替える方向
        /// </summary>
        public SortOrder Order
        {
            set
            {
                m_direction = value;
            }
            get
            {
                return m_direction;
            }
        }

        /// <summary>
        /// 並べ替え条件のクリア
        /// </summary>
        public void Clear()
        {
            m_column = 0;
            m_direction = SortOrder.None;
        }

        /// <summary>
        /// リストビューのアイテム・オブジェクトの大小関係を判定する
        /// </summary>
        /// <param name="arg1">リストビューのアイテム1</param>
        /// <param name="arg2">リストビューのアイテム2</param>
        /// <returns>各アイテムの比較結果。大小関係を返す。等しい場合は0。</returns>
        public int Compare(object item1, object item2)
        {
            int result = 0;
            switch (m_direction)
            {
                case SortOrder.None:
                    break;
                default:
                    if (null != item1 && null != item2)
                    {
                        ListViewItem
                            enter = (ListViewItem)item1,
                            leave = (ListViewItem)item2;
                        string value = enter.SubItems[m_column].Text;
                        result = value.CompareTo(leave.SubItems[m_column].Text);
                        switch (m_direction)
                        {
                            case SortOrder.Descending:
                                result = -result;
                                break;
                        }
                    }
                    break;
            }
            return result;
        }
    }
}
