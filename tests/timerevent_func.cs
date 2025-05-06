using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace tests
{
    internal class timerevent_func
    {
        void MyTimerMethod(object sender, EventArgs e)
        {
            // タイマメソッド
            // this.TextBlock1.Text = $"{DateTime.Now.ToString("HH:mm:ss")}{eventtext}";
            Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")}");


            //入力されているlotが既存のものの場合、継続工程としてフラグを立てる
            continueLotNumFlag = false;
            foreach (string lotlistdata in LotNumList)
            {
                if (lotlistdata == LotItem)
                {
                    continueLotNumFlag = true;// 一致しているので既存：true
                }
            }

            var taskAsyncFunc = Task.Run(() =>
            {
                //処理の実行
                //viewのtextboxの文字列を読み込みアクティビティを実行
                (Dt_temp, Dt_ConstructionLayout_temp) = Model_ActivitiyToSearchMpf.runActivity(
                    TextForSearch
                    , Suffixitem
                    , targetMPFversion
                    , Dt
                    , Dt_ConstructionLayout_temp
                    );

            });
            await taskAsyncFunc;

            if (Dt_temp == null && Dt_ConstructionLayout_temp == null)
            {
                //検索失敗時
                return;
            }


            DGV_release.Clear();
            DGV_release.Add(new DGV_release_DT { ReleaseDate = Dt_temp[0].ReleaseDate });

            DGV_Header_kiji.Clear();
            DGV_Header_kiji.Add(new DGV_Header_kiji_DT { 版 = Dt_temp[0].ReleaseNum, 記事 = Dt_temp[0].Article, 承認 = Dt_temp[0].ver_Authorizer });
            DGV_Header_kiji2col.Clear();
            DGV_Header_kiji2col.Add(new DGV_Header_kiji2col_DT { 版 = Dt_temp[0].ReleaseNum2, 記事 = Dt_temp[0].Article2, 承認 = Dt_temp[0].ver_Authorizer2 });
            DGV_Header_itemname.Clear();
            DGV_Header_itemname.Add(new DGV_Header_itemname_DT { MPF図仕番号 = Dt_temp[0].MPF_Number, 製品名 = Dt_temp[0].ProductName });
            DGV_Header_reviwernames.Clear();
            DGV_Header_reviwernames.Add(new DGV_Header_reviwernames_DT { 担当 = Dt_temp[0].WorkernName, 査閲 = Dt_temp[0].Reviewer, 承認 = Dt_temp[0].Authorizer, E = Dt_temp[0].E_Manager, M = Dt_temp[0].M_Manager, QC確認 = Dt_temp[0].QC_Adviser });

            Dt = Dt_temp;
            //todo ★行ごとのバインドじゃないと各行の列サイズ変えられなそうなので改修予定
            {
                Dt_ConstructionLayout.Clear();
                foreach (ConstructionLayout Dt_temp in Dt_ConstructionLayout_temp)
                {
                    int _EcoContainInfo = 0;
                    if (Dt_temp.Process.ToString().Contains("<eco")
                        || Dt_temp.Symbol.ToString().Contains("<eco")
                        || Dt_temp.WorkConstruction_Specification.ToString().Contains("<eco")
                        || Dt_temp.MaterialName_DesignConstruction.ToString().Contains("<eco")
                        || Dt_temp.Conditions_Notes.ToString().Contains("<eco")
                        )
                    {
                        _EcoContainInfo = 1;
                    }



                    Dt_ConstructionLayout.Add(new DGV_layout_DT
                    {
                        工程 = Dt_temp.Process
                    ,
                        記号 = Dt_temp.Symbol
                    ,
                        作業_工事仕様書 = Dt_temp.WorkConstruction_Specification
                    ,
                        資材名_指定 = Dt_temp.MaterialName_DesignConstruction
                    ,
                        工事条件_備考 = Dt_temp.Conditions_Notes
                        ,
                        EcoContainInfo = _EcoContainInfo
                    });
                }
            }





            //関連図示　追加：ECO読取り比較
            string af_kumizu_ver = "";
            string af_e_buhin_ver = "";
            string af_git_ver = "";
            string af_kairozu_ver = "";
            string af_m_buhin_ver = "";





            {
                {//版数比較
                    ObservableCollection<AppData.DocMasterTableClass> read_docs = new ObservableCollection<AppData.DocMasterTableClass>();
                    {
                        string target_name = $"{kumizu_num}";
                        AccessBI_SYS accessBI_SYS = new AccessBI_SYS();
                        string query = "";
                        query = $"" +
                                        $"select * from {InfoDataBase_default.db_bisys_table_docmaster} " +
                                        $"where {InfoDataBase_default.db_bisys_col_docnum} = '{target_name}';";
                        read_docs = accessBI_SYS.getDocMasterTable(query);
                    }
                    System.Collections.Generic.List<string> tempverlist = new System.Collections.Generic.List<string>();
                    tempverlist.Clear();
                    foreach (AppData.DocMasterTableClass read_doc in read_docs)
                    {
                        tempverlist.Add(read_doc.版数.ToString());
                    }

                    string pretarget_ver = "";
                    foreach (string tempver in tempverlist)
                    {
                        //最新版の情報を格納
                        if (pretarget_ver == "")
                        {
                            pretarget_ver = tempver;
                        }
                        else if (int.Parse(tempver.Substring(tempver.LastIndexOf('_') + 1)) > int.Parse(pretarget_ver.Substring(pretarget_ver.LastIndexOf('_') + 1)))
                        {
                            pretarget_ver = tempver;
                        }
                    }
                    af_kumizu_ver = pretarget_ver;
                }
                {//版数比較
                    ObservableCollection<AppData.DocMasterTableClass> read_docs = new ObservableCollection<AppData.DocMasterTableClass>();
                    {
                        string target_name = $"{e_buhin_num}";
                        AccessBI_SYS accessBI_SYS = new AccessBI_SYS();
                        string query = "";
                        query = $"" +
                                        $"select * from {InfoDataBase_default.db_bisys_table_docmaster} " +
                                        $"where {InfoDataBase_default.db_bisys_col_docnum} = '{target_name}';";
                        read_docs = accessBI_SYS.getDocMasterTable(query);
                    }
                    System.Collections.Generic.List<string> tempverlist = new System.Collections.Generic.List<string>();
                    tempverlist.Clear();
                    foreach (AppData.DocMasterTableClass read_doc in read_docs)
                    {
                        tempverlist.Add(read_doc.版数.ToString());
                    }

                    string pretarget_ver = "";
                    foreach (string tempver in tempverlist)
                    {
                        //最新版の情報を格納
                        if (pretarget_ver == "")
                        {
                            pretarget_ver = tempver;
                        }
                        else if (int.Parse(tempver.Substring(tempver.LastIndexOf('_') + 1)) > int.Parse(pretarget_ver.Substring(pretarget_ver.LastIndexOf('_') + 1)))
                        {
                            pretarget_ver = tempver;
                        }
                    }
                    af_e_buhin_ver = pretarget_ver;
                }
                {//版数比較
                    ObservableCollection<AppData.DocMasterTableClass> read_docs = new ObservableCollection<AppData.DocMasterTableClass>();
                    {
                        string target_name = $"{git_num}";
                        AccessBI_SYS accessBI_SYS = new AccessBI_SYS();
                        string query = "";
                        query = $"" +
                                        $"select * from {InfoDataBase_default.db_bisys_table_docmaster} " +
                                        $"where {InfoDataBase_default.db_bisys_col_docnum} = '{target_name}';";
                        read_docs = accessBI_SYS.getDocMasterTable(query);
                    }
                    System.Collections.Generic.List<string> tempverlist = new System.Collections.Generic.List<string>();
                    tempverlist.Clear();
                    foreach (AppData.DocMasterTableClass read_doc in read_docs)
                    {
                        tempverlist.Add(read_doc.版数.ToString());
                    }

                    string pretarget_ver = "";
                    foreach (string tempver in tempverlist)
                    {
                        //最新版の情報を格納
                        if (pretarget_ver == "")
                        {
                            pretarget_ver = tempver;
                        }
                        else if (int.Parse(tempver.Substring(tempver.LastIndexOf('_') + 1)) > int.Parse(pretarget_ver.Substring(pretarget_ver.LastIndexOf('_') + 1)))
                        {
                            pretarget_ver = tempver;
                        }
                    }
                    af_git_ver = pretarget_ver;
                }
                {//版数比較
                    ObservableCollection<AppData.DocMasterTableClass> read_docs = new ObservableCollection<AppData.DocMasterTableClass>();
                    {
                        string target_name = $"{kairozu_num}";
                        AccessBI_SYS accessBI_SYS = new AccessBI_SYS();
                        string query = "";
                        query = $"" +
                                        $"select * from {InfoDataBase_default.db_bisys_table_docmaster} " +
                                        $"where {InfoDataBase_default.db_bisys_col_docnum} = '{target_name}';";
                        read_docs = accessBI_SYS.getDocMasterTable(query);
                    }
                    System.Collections.Generic.List<string> tempverlist = new System.Collections.Generic.List<string>();
                    tempverlist.Clear();
                    foreach (AppData.DocMasterTableClass read_doc in read_docs)
                    {
                        tempverlist.Add(read_doc.版数.ToString());
                    }

                    string pretarget_ver = "";
                    foreach (string tempver in tempverlist)
                    {
                        //最新版の情報を格納
                        if (pretarget_ver == "")
                        {
                            pretarget_ver = tempver;
                        }
                        else if (int.Parse(tempver.Substring(tempver.LastIndexOf('_') + 1)) > int.Parse(pretarget_ver.Substring(pretarget_ver.LastIndexOf('_') + 1)))
                        {
                            pretarget_ver = tempver;
                        }
                    }
                    kairozu_ver = pretarget_ver;
                }
                {//版数比較
                    ObservableCollection<AppData.DocMasterTableClass> read_docs = new ObservableCollection<AppData.DocMasterTableClass>();
                    {
                        string target_name = $"{m_buhin_num}";
                        AccessBI_SYS accessBI_SYS = new AccessBI_SYS();
                        string query = "";
                        query = $"" +
                                        $"select * from {InfoDataBase_default.db_bisys_table_docmaster} " +
                                        $"where {InfoDataBase_default.db_bisys_col_docnum} = '{target_name}';";
                        read_docs = accessBI_SYS.getDocMasterTable(query);
                    }
                    System.Collections.Generic.List<string> tempverlist = new System.Collections.Generic.List<string>();
                    tempverlist.Clear();
                    foreach (AppData.DocMasterTableClass read_doc in read_docs)
                    {
                        tempverlist.Add(read_doc.版数.ToString());
                    }

                    string pretarget_ver = "";
                    foreach (string tempver in tempverlist)
                    {
                        //最新版の情報を格納
                        if (pretarget_ver == "")
                        {
                            pretarget_ver = tempver;
                        }
                        else if (int.Parse(tempver.Substring(tempver.LastIndexOf('_') + 1)) > int.Parse(pretarget_ver.Substring(pretarget_ver.LastIndexOf('_') + 1)))
                        {
                            pretarget_ver = tempver;
                        }
                    }
                    m_buhin_ver = pretarget_ver;
                }
            }

            //コンフィギュレーション比較
            {
                //コンフィギュレーションが同じ場合、ウィンドウ表示
                if (
                       (af_kumizu_ver != kumizu_ver)
                    || (af_e_buhin_ver != e_buhin_ver)
                    || (af_git_ver != git_ver)
                    || (af_kairozu_ver != kairozu_ver)
                    || (af_m_buhin_ver != m_buhin_ver)
                    ) {
                    ///window show configuration aleart


                }
            }

        }
    }
}
