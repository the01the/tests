using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace tests.Logger
{
    /// <summary>
    /// 3\. データモデル (`MyDataItem.cs`)
    ///キャッシュするデータの型を定義します。
    /// </summary>
    public class MyDataItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Value { get; set; }
        public DateTime LastUpdated { get; set; } // キャッシュデータの鮮度管理に役立つ
    }
}
