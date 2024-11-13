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

//���\�t�g�ɂ�EPPlus4.5.3.3���g�p���Ă��܂��B�����LGPL���C�Z���X�ł��B���쌠��EPPlus Software�Ђł��B

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

            //�d�C���f�����瓮�́A�R���Z���g�d����S�Ē��o����
            //�@�B���f������@�B�ݔ��@���S�Ē��o���āA�d����ʁA����d�͂Ȃǒ��o
            //�d�C�Ƌ@�B�ŁA�@��ԍ��A�@�햼�A�d�C��ʁA�e�ʂ��r����B
            //������Ԃ𔻒肵�s�����ӏ������F�Z���ɂ���Excel�ɏo�͂���

            //����***************
            //�d�C�ݔ��őΏۂƂ���[�t�@�~����,�V���{����]
            string[] elecTarget = new string[] { "20301_���t�R���Z���g_�I�o�^_200V", "20308_�W���C���g�{�b�N�X_���d�p" };
            //�d�C�Ώۗv�f�̃��X�g
            List<ElecEquip> elecInstances = new List<ElecEquip>();
            //�@�B�ݔ��őΏۂƂ���t�@�~�����ƃV���{����
            string[] mechTarget = new string[] { "07052_PAC-CK4_�����@_�J�Z�b�g�`(4����)", "07062_ACP-FRV(J)_�����@_���u(�I�o)���`(����)" };
            //�@�B�Ώۗv�f�̃��X�g
            List<MechEquip> mechInstances = new List<MechEquip>();
            //*******************


            //�d�C�C���X�^���X��S�Ē��o����
            FilteredElementCollector elecCol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_ElectricalFixtures);
            //�d�C�̑Ώۃt�@�~���A�^�C�v���i��
            for (int i = 0; i < elecTarget.Length; i++)
            {
                //Querry�ŗv�f�i��
                var query = from element in elecCol.Cast<FamilyInstance>()
                            where (element.Symbol.Family.Name == elecTarget[i])
                            select element;
                //���o�������̂����X�g�ɂ���
                List<FamilyInstance> elecFamilyInstances = query.ToList<FamilyInstance>();
                foreach(FamilyInstance elecInstance in elecFamilyInstances)
                {
                    ElecEquip eEquip = new ElecEquip(elecInstance);
                    elecInstances.Add(eEquip);
                }
            }


            //�@�B�ݔ��@���S�Ē��o����
            //�C���X�^���X��S�Ē��o����
            FilteredElementCollector mechCol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment);
            //�@�B�̑Ώۃt�@�~���A�^�C�v���i��
            for (int i = 0; i < mechTarget.Length; i++)
            {
                //Querry�ŗv�f�i��
                var query = from element in mechCol.Cast<FamilyInstance>()
                            where (element.Symbol.Family.Name == mechTarget[i])
                            select element;
                //���o�������̂����X�g�ɂ���
                List<FamilyInstance> mechFamilyInstances = query.ToList<FamilyInstance>();
                foreach (FamilyInstance mechInstance in mechFamilyInstances)
                {
                    MechEquip mEquip = new MechEquip(mechInstance);
                    mechInstances.Add(mEquip);
                }
            }


            //�d�C�Ƌ@�B�ŁA�@��ԍ��A�@�햼�A�d�C��ʁA�e�ʂ��r����B

            //�u�@��L���v���L�[�Ƃ��ċ@�B�Ɠd�C�̑Ή��֌W���m�F����
            //�\�`���̃f�[�^���������߂�DataTable���쐬����
            var datatable = new DataTable("tblData");
            //�񖼁A�^��ݒ肷��
            datatable.Columns.Add("M �L��", typeof(string));//0
            datatable.Columns.Add("M �t�@�~����", typeof(string));//1
            datatable.Columns.Add("M �^�C�v��", typeof(string));//2
            datatable.Columns.Add("M �ɐ�", typeof(int));//3
            datatable.Columns.Add("M ��", typeof(int));//4
            datatable.Columns.Add("M �d��", typeof(double));//5
            datatable.Columns.Add("M ����d��_��[", typeof(double));//6
            datatable.Columns.Add("M ����d��_�g�[", typeof(double));//7
            datatable.Columns.Add("M �ʒu���W", typeof(string));//8


            datatable.Columns.Add("E �L��", typeof(string));//9
            datatable.Columns.Add("E �t�@�~����", typeof(string));//10
            datatable.Columns.Add("E �^�C�v��", typeof(string));//11
            datatable.Columns.Add("E �ɐ�", typeof(int));//12
            datatable.Columns.Add("E ��", typeof(int));//13
            datatable.Columns.Add("E �d��", typeof(double));//14
            datatable.Columns.Add("E �瑊�d��", typeof(double));//15
            datatable.Columns.Add("E �ʒu���W", typeof(string));//16

            datatable.Columns.Add("����mm", typeof(double));//17


            //�@�탊�X�g����e�[�u������
            foreach (MechEquip eqp in mechInstances)
            {
                var row = datatable.NewRow();

                //�@�B�ݔ��̃f�[�^���́@�A�p�����[�^�����݂��Ȃ��ꍇ�̃G���[�������ƍl������K�v����
                if (eqp.Kigou != null) { row[0] = eqp.Kigou; } else { row[0] = DBNull.Value; }
                if (eqp.Family != null) { row[1] = eqp.Family; } else { row[1] = DBNull.Value; }
                if (eqp.Type != null) { row[2] = eqp.Type; } else { row[2] = DBNull.Value; }
                if (eqp.Pole != null) { row[3] = eqp.Pole; } else { row[3] = DBNull.Value; }
                if (eqp.Phase != null) { row[4] = eqp.Phase; } else { row[4] = DBNull.Value; }
                row[5] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.Voltage), UnitTypeId.Volts), 0);
                row[6] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.ReibouShouhidenryoku), UnitTypeId.Watts), 0);
                row[7] = Math.Round(UnitUtils.ConvertFromInternalUnits(KUtil.tryDouble(eqp.DanbouShouhidenryoku), UnitTypeId.Watts), 0);
                row[8] = eqp.Loation.ToString();

                //�Ή�����d�C����������f�[�^�����A�p�����[�^�����݂��Ȃ��ꍇ�̃G���[�������ƍl������K�v����A�Ή��Ȃ����̂̏����l���K�v
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
                            row[17] = Distance(row[8].ToString(), row[16].ToString());//����
                        }
                    }
                }
                datatable.Rows.Add(row);
            }

            

            //���ɁA���ꗗ��Excel�ɕۑ����邽�߂̃_�C�A���O���J��
            System.Windows.Forms.SaveFileDialog saveFileDialog = KUtil.SaveExcel();
            String filename = saveFileDialog.FileName;

            //�ۑ��t�@�C����������ɓ��͂���Ă����Excel�ۑ����s
            if (filename != "")
            {
                //���[�U�[���w�肵��Excel�t�@�C�����Ńt�@�C���X�g���[�����擾����
                FileStream fs = (FileStream)saveFileDialog.OpenFile();
                //ExcelPackage���쐬����
                using (ExcelPackage package = new ExcelPackage(fs))
                {
                    //Excel�̃��[�N�V�[�g��Rooms�Ƃ������̂̃V�[�g�����
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Seigousei");
                    //Excel�̈�ԍ���̈ʒuA1����DataTable����C�ɗ�������
                    worksheet.Cells["A1"].LoadFromDataTable(datatable, true);

                    //�f�[�^�͈̔͂��擾����
                    int maxRow = worksheet.Dimension.Rows;
                    int maxColomun = worksheet.Dimension.Columns;


                    //���������`�F�b�N����A�Ⴄ�Ƃ���̓Z�������F�ɂ���
                    for (int i = 2; i <= maxRow; i++)//�C���f�b�N�X�ɒ��Ӂ@�f�[�^��2�s�ڂ��瑶�݂���
                    {
                        //Pole
                        if (worksheet.Cells[i, 4].Value.ToString() != worksheet.Cells[i, 13].Value.ToString())
                        {
                            CellColor(worksheet, i, 13, System.Drawing.Color.Yellow);
                        }
                        //��

                        if (worksheet.Cells[i, 5].Value != worksheet.Cells[i, 14].Value)
                        {
                            CellColor(worksheet, i, 14, System.Drawing.Color.Yellow);
                        }

                        //�d��
                        if (worksheet.Cells[i, 6].Value.ToString() != worksheet.Cells[i, 15].Value.ToString())
                        {
                            CellColor(worksheet, i, 15, System.Drawing.Color.Yellow);
                        }
                        //�d��
                        double reibouWatt = Convert.ToDouble(worksheet.Cells[i, 7].Value.ToString());
                        double danbouWatt = Convert.ToDouble(worksheet.Cells[i, 8].Value.ToString());
                        double denkiShouhi = Convert.ToDouble(worksheet.Cells[i, 16].Value.ToString());

                        if (Math.Max(reibouWatt, danbouWatt) != denkiShouhi)
                        {
                            CellColor(worksheet, i, 16, System.Drawing.Color.Yellow);
                        }
                        //����
                        if (Convert.ToDouble(worksheet.Cells[i, 18].Value.ToString()) > 3000 ) 
                        {
                            CellColor(worksheet, i, 18, System.Drawing.Color.Yellow);
                        }
                    }

                    //�񕝂̓I�[�g�t�B�b�g������B�Ō�Ɏ��s�����ق����悢
                    for (int i = 0; i < maxColomun; i++)
                    {
                        worksheet.Column(i + 1).AutoFit();
                        worksheet.Column(i + 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }


                    //Excel�ҏW�̏I���@Excel�t�@�C����ۑ�����
                    package.Save();

                    //���펞�Ƀ��b�Z�[�W���o��
                    TaskDialog.Show("����I��", "Excel�o�͂͐���I���B���̌�Excel�t�@�C�����J���܂��B");
                    //�m�F�̂��߂�Excel���N�����ĕۑ������t�@�C�����J��
                    Process process = new Process();
                    process.StartInfo.FileName = "excel.exe";
                    process.StartInfo.Arguments = saveFileDialog.FileName;
                    process.Start();
                }
            }

                    return Result.Succeeded;
        }
        //Excel�̃Z���ɒ��F����
        private void CellColor(ExcelWorksheet sheet, int row, int col, System.Drawing.Color color)
        {
            sheet.Cells[row , col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row , col].Style.Fill.BackgroundColor.SetColor(color);
        }
        //2�_�Ԃ̋������v�Z����
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
            distance = UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters);//�~�����[�^�[
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
            //�p�����[�^�͒l�������ꍇ�ɂ�null�ɂȂ�

            //�t�@�~���ƃ^�C�v
            FamilySymbol familySymbol = instance.Symbol;
            Family family = familySymbol.Family;


            //�C���X�^���X
            this.Instance = instance;
            //���݂̃G�������g�̈ʒu
            this.Loation = KUtil.GetElementLocation(instance);
            //�@�B�ݔ���KM�@��L��
            this.Kigou = KUtil.ShowParameter(instance, "�L��");
            //�t�@�~����
            this.Family = family.Name;
            //�^�C�v��
            this.Type = familySymbol.Name;
            //�ȉ��̓t�@�~���V���{���i�^�C�v�j���璊�o����i�^�C�v�p�����[�^�̂��߁j
            //�ɐ�
            this.Pole = KUtil.ShowParameter(familySymbol, "�ɐ�");
            //��
            this.Phase = KUtil.ShowParameter(familySymbol, "��");
            //�d��
            this.Voltage = KUtil.ShowParameter(familySymbol, "�d��");
            //�@�B�ݔ��̓d�C�@����d�́@��[��
            this.ReibouShouhidenryoku = KUtil.ShowParameter(familySymbol, "����d��_��[");
            //�@�B�ݔ��̓d�C�@����d�́@�g�[��
            this.DanbouShouhidenryoku = KUtil.ShowParameter(familySymbol, "����d��_�g�[");
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
            //�p�����[�^�͒l�������ꍇ�ɂ�null�ɂȂ�

            //�t�@�~���ƃ^�C�v
            FamilySymbol familySymbol = instance.Symbol;
            Family family = familySymbol.Family;

            //�C���X�^���X
            this.Instance = instance;
            //���݂̃G�������g�̈ʒu
            this.Loation = KUtil.GetElementLocation(instance);
            //�d�C�ݔ���KM�@��L��
            this.Kigou = KUtil.ShowParameter(instance, "�L��");
            //�t�@�~����
            this.Family = family.Name;
            //�^�C�v��
            this.Type = familySymbol.Name;
            //�d�C�ݔ��̓d�C�@�ɐ�
            this.Pole = KUtil.ShowParameter(familySymbol, "�ɐ�");
            //�d�C�ݔ��̑�
            this.Phase = KUtil.ShowParameter(familySymbol, "��");
            //�d�C�ݔ��̓d��
            this.Voltage = KUtil.ShowParameter(familySymbol, "�d��");
            //�d�C�ݔ��̏���d�́i�瑊�d�́j
            this.Hisoudenryoku = KUtil.ShowParameter(instance, "�瑊�d��");
        }
    }
}
