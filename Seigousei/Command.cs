#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using RoomInfo;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;

#endregion

//当ソフトにはEPPlus4.5.3.3を使用しています。これはLGPLライセンスです。著作権はEPPlus Software社です。

namespace Seigousei
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //電気モデルから動力、コンセント電源を全て抽出する
            //機械モデルから機械設備機器を全て抽出して、電源種別、消費電力など抽出
            //電気と機械で、機器番号、機器名、電気種別、容量を比較する。
            //整合状態を判定し不整合箇所を黄色セルにしてExcelに出力する

            //条件***************
            //電気設備で対象とする[ファミリ名,シンボル名]
            string[] elecTarget = new string[] { "20301_鍵付コンセント_露出型_200V", "20308_ジョイントボックス_強電用" };
            //電気対象要素のリスト
            List<ElecEquip> elecInstances = new List<ElecEquip>();
            //機械設備で対象とするファミリ名とシンボル名
            string[] mechTarget = new string[] { "07052_PAC-CK4_室内機_カセット形(4方向)", "07062_ACP-FRV(J)_室内機_床置(露出)立形(直吹)" };
            //機械対象要素のリスト
            List<MechEquip> mechInstances = new List<MechEquip>();
            //*******************


            //電気インスタンスを全て抽出する
            FilteredElementCollector elecCol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_ElectricalFixtures);
            //電気の対象ファミリ、タイプを絞る
            for (int i = 0; i < elecTarget.Length; i++)
            {
                //Querryで要素絞る
                var query = from element in elecCol.Cast<FamilyInstance>()
                            where (element.Symbol.Family.Name == elecTarget[i])
                            select element;
                //抽出したものをリストにする
                List<FamilyInstance> elecFamilyInstances = query.ToList<FamilyInstance>();
                foreach(FamilyInstance elecInstance in elecFamilyInstances)
                {
                    ElecEquip eEquip = new ElecEquip(elecInstance);
                    elecInstances.Add(eEquip);
                }
            }


            //機械設備機器を全て抽出する
            //インスタンスを全て抽出する
            FilteredElementCollector mechCol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment);
            //機械の対象ファミリ、タイプを絞る
            for (int i = 0; i < mechTarget.Length; i++)
            {
                //Querryで要素絞る
                var query = from element in mechCol.Cast<FamilyInstance>()
                            where (element.Symbol.Family.Name == mechTarget[i])
                            select element;
                //抽出したものをリストにする
                List<FamilyInstance> mechFamilyInstances = query.ToList<FamilyInstance>();
                foreach (FamilyInstance mechInstance in mechFamilyInstances)
                {
                    MechEquip mEquip = new MechEquip(mechInstance);
                    mechInstances.Add(mEquip);
                }
            }


            //電気と機械で、機器番号、機器名、電気種別、容量を比較する。

            //「機器記号」をキーとして機械と電気の対応関係を確認する
            //表形式のデータを扱うためにDataTableを作成する
            var datatable = new DataTable("tblData");
            //列名、型を設定する
            datatable.Columns.Add("M 記号", typeof(string));//0
            datatable.Columns.Add("M ファミリ名", typeof(string));//1
            datatable.Columns.Add("M タイプ名", typeof(string));//2
            datatable.Columns.Add("M 極数", typeof(int));//3
            datatable.Columns.Add("M 相", typeof(int));//4
            datatable.Columns.Add("M 電圧", typeof(double));//5
            datatable.Columns.Add("M 消費電力_冷房", typeof(double));//6
            datatable.Columns.Add("M 消費電力_暖房", typeof(double));//7
            datatable.Columns.Add("M 位置座標", typeof(string));//8


            datatable.Columns.Add("E 記号", typeof(string));//9
            datatable.Columns.Add("E ファミリ名", typeof(string));//10
            datatable.Columns.Add("E タイプ名", typeof(string));//11
            datatable.Columns.Add("E 極数", typeof(int));//12
            datatable.Columns.Add("E 相", typeof(int));//13
            datatable.Columns.Add("E 電圧", typeof(double));//14
            datatable.Columns.Add("E 皮相電力", typeof(double));//15
            datatable.Columns.Add("E 位置座標", typeof(string));//16

            datatable.Columns.Add("距離mm", typeof(double));//17


            //機器リストからテーブル入力
            foreach (MechEquip eqp in mechInstances)
            {
                var row = datatable.NewRow();

                //機械設備のデータ入力　、パラメータが存在しない場合のエラーをもっと考慮する必要ある
                if (eqp.Kigou != null) { row[0] = eqp.Kigou; } else { row[0] = DBNull.Value; }
                if (eqp.Family != null) { row[1] = eqp.Family; } else { row[1] = DBNull.Value; }
                if (eqp.Type != null) { row[2] = eqp.Type; } else { row[2] = DBNull.Value; }
                if (eqp.Pole != null) { row[3] = eqp.Pole; } else { row[3] = DBNull.Value; }
                if (eqp.Phase != null) { row[4] = eqp.Phase; } else { row[4] = DBNull.Value; }
                row[5] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.Voltage), UnitTypeId.Volts), 0);
                row[6] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.ReibouShouhidenryoku), UnitTypeId.Watts), 0);
                row[7] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.DanbouShouhidenryoku), UnitTypeId.Watts), 0);
                row[8] = eqp.Loation.ToString();

                //対応する電気があったらデータを代入、パラメータが存在しない場合のエラーをもっと考慮する必要ある、対応ないものの処理考慮必要
                foreach (ElecEquip elecEquip in elecInstances)
                {
                    if(elecEquip.Kigou != null)
                    {
                        if (elecEquip.Kigou == eqp.Kigou)
                        {
                            if (elecEquip.Kigou != null) { row[9] = elecEquip.Kigou; } else { row[9] = DBNull.Value; }
                            if (elecEquip.Family != null) { row[10] = elecEquip.Family; } else { row[10] = DBNull.Value; }
                            if (elecEquip.Type != null) { row[11] = elecEquip.Type; } else { row[11] = DBNull.Value; }
                            if (elecEquip.Pole != null) { row[12] = elecEquip.Pole; } else { row[12] = DBNull.Value; }
                            if (elecEquip.Phase != null) { row[13] = elecEquip.Phase; } else { row[13] = DBNull.Value; }
                            row[14] = UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(elecEquip.Voltage), UnitTypeId.Volts);
                            row[15] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(elecEquip.Hisoudenryoku), UnitTypeId.Watts), 0);
                            row[16] = elecEquip.Loation.ToString();
                            row[17] = Distance(row[8].ToString(), row[16].ToString());//距離
                        }
                    }
                }
                datatable.Rows.Add(row);
            }

            

            //次に、室一覧をExcelに保存するためのダイアログを開く
            System.Windows.Forms.SaveFileDialog saveFileDialog = KUtil.SaveExcel();
            String filename = saveFileDialog.FileName;

            //保存ファイル名が正常に入力されていればExcel保存実行
            if (filename != "")
            {
                //ユーザーが指定したExcelファイル名でファイルストリームを取得する
                FileStream fs = (FileStream)saveFileDialog.OpenFile();
                //ExcelPackageを作成する
                using (ExcelPackage package = new ExcelPackage(fs))
                {
                    //ExcelのワークシートにRoomsという名称のシートを作る
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Seigousei");
                    //Excelの一番左上の位置A1からDataTableを一気に流し込む
                    worksheet.Cells["A1"].LoadFromDataTable(datatable, true);

                    //データの範囲を取得する
                    int maxRow = worksheet.Dimension.Rows;
                    int maxColomun = worksheet.Dimension.Columns;


                    //整合性をチェックする、違うところはセルを黄色にする
                    for (int i = 2; i <= maxRow; i++)//インデックスに注意　データは2行目から存在する
                    {
                        //Pole
                        if (worksheet.Cells[i, 4].Value.ToString() != worksheet.Cells[i, 13].Value.ToString())
                        {
                            CellColor(worksheet, i, 13, System.Drawing.Color.Yellow);
                        }
                        //相

                        if (worksheet.Cells[i, 5].Value != worksheet.Cells[i, 14].Value)
                        {
                            CellColor(worksheet, i, 14, System.Drawing.Color.Yellow);
                        }

                        //電圧
                        if (worksheet.Cells[i, 6].Value.ToString() != worksheet.Cells[i, 15].Value.ToString())
                        {
                            CellColor(worksheet, i, 15, System.Drawing.Color.Yellow);
                        }
                        //電力
                        double reibouWatt = Convert.ToDouble(worksheet.Cells[i, 7].Value.ToString());
                        double danbouWatt = Convert.ToDouble(worksheet.Cells[i, 8].Value.ToString());
                        double denkiShouhi = Convert.ToDouble(worksheet.Cells[i, 16].Value.ToString());

                        if (Math.Max(reibouWatt, danbouWatt) != denkiShouhi)
                        {
                            CellColor(worksheet, i, 16, System.Drawing.Color.Yellow);
                        }
                        //距離
                        if (Convert.ToDouble(worksheet.Cells[i, 18].Value.ToString()) > 3000 ) 
                        {
                            CellColor(worksheet, i, 18, System.Drawing.Color.Yellow);
                        }
                    }

                    //列幅はオートフィットさせる。最後に実行したほうがよい
                    for (int i = 0; i < maxColomun; i++)
                    {
                        worksheet.Column(i + 1).AutoFit();
                        worksheet.Column(i + 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }


                    //Excel編集の終了　Excelファイルを保存する
                    package.Save();

                    //正常時にメッセージを出す
                    TaskDialog.Show("正常終了", "Excel出力は正常終了。この後Excelファイルを開きます。");
                    //確認のためにExcelを起動して保存したファイルを開く
                    Process process = new Process();
                    process.StartInfo.FileName = "excel.exe";
                    process.StartInfo.Arguments = saveFileDialog.FileName;
                    process.Start();
                }
            }

                    return Result.Succeeded;
        }
        //Excelのセルに着色する
        private void CellColor(ExcelWorksheet sheet, int row, int col, System.Drawing.Color color)
        {
            sheet.Cells[row , col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row , col].Style.Fill.BackgroundColor.SetColor(color);
        }
        //2点間の距離を計算する
        private double Distance(string obj1, string obj2)
        {
            double distance;
            obj1 = obj1.Trim().Replace("(", "").Replace(")", "");
            obj2 = obj2.Trim().Replace("(", "").Replace(")", "");
            string[] obj11 = obj1.Split(',');
            double x1 = double.Parse(obj11[0]);double y1 = double.Parse(obj11[1]);double z1 = double.Parse(obj11[2]);
            string[] obj21 = obj2.Split(',');
            double x2 = double.Parse(obj21[0]);double y2 = double.Parse(obj21[1]);double z2= double.Parse(obj21[2]);
            distance = Math.Sqrt(Math.Abs(x1 - x2)* Math.Abs(x1 - x2) + Math.Abs(y1 - y2) * Math.Abs(y1 - y2)+ Math.Abs(z1 - z2) * Math.Abs(z1 - z2));
            distance = UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters);//ミリメーター
            distance = Math.Round(distance);
            return distance;
        }
    }
    public class MechEquip
    {
        public FamilyInstance Instance { private set; get; }
        public XYZ Loation { private set; get; }
        public string Kigou { private set; get; }
        public string Family { private set; get; }
        public string Type { private set; get; }
        public string Pole { private set; get; }
        public string Phase { private set; get; }
        public string Voltage { private set; get; }
        public string ReibouShouhidenryoku { private set; get; }
        public string DanbouShouhidenryoku { private set; get; }


        public MechEquip(FamilyInstance instance)
        {
            //パラメータは値が無い場合にはnullになる

            //ファミリとタイプ
            FamilySymbol familySymbol = instance.Symbol;
            Family family = familySymbol.Family;


            //インスタンス
            this.Instance = instance;
            //現在のエレメントの位置
            this.Loation = KUtil.GetElementLocation(instance);
            //機械設備のKM機器記号
            this.Kigou = KUtil.ShowParameter(instance, "記号");
            //ファミリ名
            this.Family = family.Name;
            //タイプ名
            this.Type = familySymbol.Name;
            //以下はファミリシンボル（タイプ）から抽出する（タイプパラメータのため）
            //極数
            this.Pole = KUtil.ShowParameter(familySymbol, "極数");
            //相
            this.Phase = KUtil.ShowParameter(familySymbol, "相");
            //電圧
            this.Voltage = KUtil.ShowParameter(familySymbol, "電圧");
            //機械設備の電気　消費電力　冷房時
            this.ReibouShouhidenryoku = KUtil.ShowParameter(familySymbol, "消費電力_冷房");
            //機械設備の電気　消費電力　暖房時
            this.DanbouShouhidenryoku = KUtil.ShowParameter(familySymbol, "消費電力_暖房");
        }
    }
    public class ElecEquip
    {
        public FamilyInstance Instance { private set; get; }
        public XYZ Loation { private set; get; }
        public string Kigou { private set; get; }
        public string Family { private set; get; }
        public string Type { private set; get; }
        public string Pole { private set; get; }
        public string Phase { private set; get; }
        public string Voltage { private set; get; }
        public string Hisoudenryoku { private set; get; }


        public ElecEquip(FamilyInstance instance)
        {
            //パラメータは値が無い場合にはnullになる

            //ファミリとタイプ
            FamilySymbol familySymbol = instance.Symbol;
            Family family = familySymbol.Family;

            //インスタンス
            this.Instance = instance;
            //現在のエレメントの位置
            this.Loation = KUtil.GetElementLocation(instance);
            //電気設備のKM機器記号
            this.Kigou = KUtil.ShowParameter(instance, "記号");
            //ファミリ名
            this.Family = family.Name;
            //タイプ名
            this.Type = familySymbol.Name;
            //電気設備の電気　極数
            this.Pole = KUtil.ShowParameter(familySymbol, "極数");
            //電気設備の相
            this.Phase = KUtil.ShowParameter(familySymbol, "相");
            //電気設備の電圧
            this.Voltage = KUtil.ShowParameter(familySymbol, "電圧");
            //電気設備の消費電力（皮相電力）
            this.Hisoudenryoku = KUtil.ShowParameter(instance, "皮相電力");
        }
    }
}
