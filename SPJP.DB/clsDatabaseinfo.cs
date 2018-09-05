using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
//using System.Threading.Tasks;

namespace SPJP.DB
{
    public class clsOrderDatabaseinfo
    {
        public string Order_id { get; set; }
        public string patientId { get; set; }//患者ID
        public string patientName { get; set; }//患者姓名
        public string sex { get; set; }// 患者性别（1男,2女）
        public string birthday { get; set; }//出生年月
        public string diseaseType { get; set; }//诊断（1帕金森综合征,2.特发性震颤…）
        public string courseOfDisease { get; set; }// 病程（年）
        public string serialNumber { get; set; }//检查订单流水号
        public string inspectId { get; set; }//检查ID
        public string checkType { get; set; }//检查内容（1计时起身-步行,2窄道）
        public string dualTask { get; set; }//双任务（0否,1是）
        public string checkRemark { get; set; }//检查备注
        public string checkStartTime { get; set; }//检查开始时间
        public string checkEndTime { get; set; }//检查结束时间
        public string checkHospitalId { get; set; }//检查医院id
        public string checkHospitalName { get; set; }//检查医院
        public string checkDoctorId { get; set; }//检查医生id
        public string checkDoctorName { get; set; }//检查医生
        public string auditDoctorName { get; set; }// 审核医生
        public string checkConclusion { get; set; }//检查结论
        public string sistDuration { get; set; }//起立过程时间SiSt Duration（sec）
        public string peakSiStAngularVelocity { get; set; }//起立过程中躯干矢状面峰值角速度
        public string sistTrunkRoM { get; set; }//起立过程中躯干矢状面角度范围
        public string stepLength_R { get; set; }//右脚步长Step Length R.（cm）
        public string stepLength_L { get; set; }//左脚步长Step Length L.（cm）
        public string stepLength_avg { get; set; }//左右脚平均步长Step Length（cm）
        public string strideVelocity_R { get; set; }//右脚速率Stride Velocity R.（m/s）
        public string strideVelocity_L { get; set; }//左脚速率Stride Velocity L.（m/s）
        public string strideVelocity_avg { get; set; }//左右脚平均速率Stride Velocity（m/s）
        public string strideLength_R { get; set; }//右脚跨步长Stride Length R.（cm）
        public string strideLength_L { get; set; }//左脚跨步长Stride Length L.（cm）
        public string strideLength_avg { get; set; }//左右脚平均跨步长Stride Length（cm）
        public string gaitCycle_R { get; set; }//右脚跨步时长Gait Cycle R.（sec）
        public string gaitCycle_L { get; set; }//左脚跨步时长Gait Cycle L.（sec）
        public string gaitCycle_avg { get; set; }//左右脚平均跨步时长Gait Cycle（sec）
        public string cadence_R { get; set; }//右脚步频Cadence R.（step/min）
        public string cadence_L { get; set; }//左脚步频Cadence L.（step/min）
        public string cadence_avg { get; set; }//左右脚平均步频Cadence（step/min）     
        public string support_R { get; set; }//右腿支撑相Right Support（%）
        public string support_L { get; set; }//左腿支撑相Left Support（%）
        public string support_avg { get; set; }//双腿支撑相Double Support（%）
        public string swing_R { get; set; }//右侧摆动相Swing R.（%）
        public string swing_L { get; set; }//左侧摆动相Swing L.（%）
        public string swing_avg { get; set; }//摆动相Swing（%）
        public string stance_R { get; set; }//右侧支撑相Stance R.（%）
        public string stance_L { get; set; }//左侧支撑相Stance L.（%）
        public string stance_avg { get; set; }//支撑相Stance（%）
        public string trunkSagittalPeakVelocity { get; set; }//躯干矢状面运动峰值角速度Trunk Sagittal Peak Velocity（degree/sec）
        public string trunkHorizontalPeakVelocity { get; set; }//躯干横断面运动峰值角速度Trunk Horizontal Peak Velocity （degree/sec）
        public string strideVelocityAsymmetry { get; set; }//左右脚速率的相对偏差Stride Velocity Asymmetry（%）
        public string strideLengthAsymmetry { get; set; }//左右脚跨步长的相对偏差Stride Length Asymmetry（%）
        public string swingAsymmetry { get; set; }//左右腿摆动相相对偏差Swing Asymmetry（%）
        public string stanceAsymmetry { get; set; }//左右腿支撑相相对偏差Stance Asymmetry（%）
        public string shankAsymmetry { get; set; }//左右小腿角度范围的相对偏差Shank Asymmetry（%）
        public string peakShankVelocityAsymmetry { get; set; }//左右小腿速率的相对偏差Peak Shank Velocity Asymmetry（%）
        public string shank_SSI { get; set; }//小腿角度范围对称性指数Shank SSI  (Symbolic Symmetry Index)（%）
        public string meanPhaseDifference { get; set; }//小腿相对平均相位差Mean Phase Difference（%）
        public string phaseCoordinationIndex { get; set; }//小腿相对平均相位差Mean Phase Difference（%）
        public string balanceTrunkSagittalPeakVelocity { get; set; }//躯干矢状面峰值角速度Trunk Sagittal Peak Velocity（degree/sec）
        public string balanceTrunkHorizontalPeakVelocity { get; set; }//躯干横断面峰值角速度Trunk Horizontal Peak Velocity（degree/sec）

        //system 

        public bool showimage { get; set; }
        public string dizhi { get; set; }
        public string typeCode { get; set; }
        public string childType { get; set; }

    }


    [Serializable]
    [DataContract]
    public class result_Msg
    {
        /// <summary>
        /// code
        /// </summary>
        [DataMember(IsRequired = false)]
        public string resultCode { get; set; }
        /// <summary>
        /// msg
        /// </summary>
        [DataMember(IsRequired = false)]
        public string resultMsg { get; set; }
        /// <summary>
        /// typeCode
        /// </summary>
        [DataMember(IsRequired = false)]
        public string typeCode { get; set; }
        ///
        [DataMember(IsRequired = false)]
        public string childType { get; set; }


        /// <summary>
        /// 数据集合
        /// </summary>
        [DataMember(IsRequired = false)]
        public resultData resultData { get; set; }
    }
    [Serializable]
    public class resultData
    {
        public List<clsPDFchildTypeinfo> classinfo { get; set; }

    }

    public class clsPDFchildTypeinfo
    {
        public string typeCode { get; set; }
        public string typeName { get; set; }
        public string typeEnName { get; set; }
        public string unit { get; set; }
        public string distinguish { get; set; }

        //
        public string classNo { get; set; }
        public string className { get; set; }


    }

    public class clsExcelinfo
    {
        public string text { get; set; }
        public string size { get; set; }
        public string cols { get; set; }

        public string reqId { get; set; }

    }
    #region Excel
    public class Cols
    {
        /// <summary>
        /// 基础信息
        /// </summary>
        public string text { get; set; }
        /// <summary>
        /// Size
        /// </summary>
        public int size { get; set; }
    }

    public class Rows
    {
        /// <summary>
        /// Cols
        /// </summary>
        public List<Cols> cols { get; set; }
    }

    public class Header
    {
        /// <summary>
        /// Rows
        /// </summary>
        public List<Rows> rows { get; set; }
    }

    public class Datas
    {

        /// <summary>
        /// IndexNumber
        /// </summary>
        public int indexNumber { get; set; }
        /// <summary>
        /// PatientId
        /// </summary>
        public int patientId { get; set; }
        /// <summary>
        /// 王二
        /// </summary>
        public string patientName { get; set; }
        /// <summary>
        /// Sex
        /// </summary>
        public int sex { get; set; }
        /// <summary>
        /// 2000-01-01 00:00:00
        /// </summary>
        public DateTime birthday { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string diseaseType { get; set; }
        /// <summary>
        /// CourseOfDisease
        /// </summary>
        public int courseOfDisease { get; set; }
        /// <summary>
        /// 201808062022447132156
        /// </summary>
        public string serialNumber { get; set; }
        /// <summary>
        /// 25
        /// </summary>
        public string inspectId { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string checkType { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string dualTask { get; set; }
        /// <summary>
        /// 检查项备注内容
        /// </summary>
        public string checkRemark { get; set; }
        /// <summary>
        /// 2018-08-01 17:19:00
        /// </summary>
        public DateTime checkStartTime { get; set; }
        /// <summary>
        /// 2018-08-06 17:19:00
        /// </summary>
        public DateTime checkEndTime { get; set; }
        /// <summary>
        /// 8a9e2d385f1eb70c015f2e41a15000c5
        /// </summary>
        public string checkHospitalId { get; set; }
        /// <summary>
        /// 华侨医院
        /// </summary>
        public string checkHospitalName { get; set; }
        /// <summary>
        /// biwei
        /// </summary>
        public string checkDoctorId { get; set; }
        /// <summary>
        /// 毕伟
        /// </summary>
        public string checkDoctorName { get; set; }
        /// <summary>
        /// AuditDoctorName
        /// </summary>
        public string auditDoctorName { get; set; }
        /// <summary>
        /// 这是医生的结论
        /// </summary>
        public string checkConclusion { get; set; }
        /// <summary>
        /// 23±2
        /// </summary>
        public string sistDuration { get; set; }
        /// <summary>
        /// 22±1
        /// </summary>
        public string peakSiStAngularVelocity { get; set; }
        /// <summary>
        /// 21±0.5
        /// </summary>
        public string sistTrunkRoM { get; set; }
        /// <summary>
        /// StepLength_R
        /// </summary>
        public string stepLength_R { get; set; }
        /// <summary>
        /// StepLength_L
        /// </summary>
        public string stepLength_L { get; set; }
        /// <summary>
        /// StepLength_avg
        /// </summary>
        public string stepLength_avg { get; set; }
        /// <summary>
        /// StrideVelocity_R
        /// </summary>
        public string strideVelocity_R { get; set; }
        /// <summary>
        /// StrideVelocity_L
        /// </summary>
        public string strideVelocity_L { get; set; }
        /// <summary>
        /// StrideVelocity_avg
        /// </summary>
        public string strideVelocity_avg { get; set; }
        /// <summary>
        /// StrideLength_R
        /// </summary>
        public string strideLength_R { get; set; }
        /// <summary>
        /// StrideLength_L
        /// </summary>
        public string strideLength_L { get; set; }
        /// <summary>
        /// StrideLength_avg
        /// </summary>
        public string strideLength_avg { get; set; }
        /// <summary>
        /// GaitCycle_R
        /// </summary>
        public string gaitCycle_R { get; set; }
        /// <summary>
        /// GaitCycle_L
        /// </summary>
        public string gaitCycle_L { get; set; }
        /// <summary>
        /// GaitCycle_avg
        /// </summary>
        public string gaitCycle_avg { get; set; }
        /// <summary>
        /// Cadence_R
        /// </summary>
        public string cadence_R { get; set; }
        /// <summary>
        /// Cadence_L
        /// </summary>
        public string cadence_L { get; set; }
        /// <summary>
        /// Cadence_avg
        /// </summary>
        public string cadence_avg { get; set; }
        /// <summary>
        /// Support_R
        /// </summary>
        public string support_R { get; set; }
        /// <summary>
        /// Support_L
        /// </summary>
        public string support_L { get; set; }
        /// <summary>
        /// Support_avg
        /// </summary>
        public string support_avg { get; set; }
        /// <summary>
        /// Swing_R
        /// </summary>
        public string swing_R { get; set; }
        /// <summary>
        /// Swing_L
        /// </summary>
        public string swing_L { get; set; }
        /// <summary>
        /// Swing_avg
        /// </summary>
        public string swing_avg { get; set; }
        /// <summary>
        /// Stance_R
        /// </summary>
        public string stance_R { get; set; }
        /// <summary>
        /// Stance_L
        /// </summary>
        public string stance_L { get; set; }
        /// <summary>
        /// Stance_avg
        /// </summary>
        public string stance_avg { get; set; }
        /// <summary>
        /// TrunkSagittalPeakVelocity
        /// </summary>
        public string trunkSagittalPeakVelocity { get; set; }
        /// <summary>
        /// TrunkHorizontalPeakVelocity
        /// </summary>
        public string trunkHorizontalPeakVelocity { get; set; }
        /// <summary>
        /// StrideVelocityAsymmetry
        /// </summary>
        public string strideVelocityAsymmetry { get; set; }
        /// <summary>
        /// StrideLengthAsymmetry
        /// </summary>
        public string strideLengthAsymmetry { get; set; }
        /// <summary>
        /// SwingAsymmetry
        /// </summary>
        public string swingAsymmetry { get; set; }
        /// <summary>
        /// StanceAsymmetry
        /// </summary>
        public string stanceAsymmetry { get; set; }
        /// <summary>
        /// ShankAsymmetry
        /// </summary>
        public string shankAsymmetry { get; set; }
        /// <summary>
        /// PeakShankVelocityAsymmetry
        /// </summary>
        public string peakShankVelocityAsymmetry { get; set; }
        /// <summary>
        /// Shank_SSI
        /// </summary>
        public string shank_SSI { get; set; }
        /// <summary>
        /// MeanPhaseDifference
        /// </summary>
        public string meanPhaseDifference { get; set; }
        /// <summary>
        /// PhaseCoordinationIndex
        /// </summary>
        public string phaseCoordinationIndex { get; set; }
        /// <summary>
        /// BalanceTrunkSagittalPeakVelocity
        /// </summary>
        public string balanceTrunkSagittalPeakVelocity { get; set; }
        /// <summary>
        /// BalanceTrunkHorizontalPeakVelocity
        /// </summary>
        public string balanceTrunkHorizontalPeakVelocity { get; set; }

        public string reqId { get; set; }
    }

    public class Content
    {
        /// <summary>
        /// Datas
        /// </summary>
        public List<Datas> datas { get; set; }
    }

    public class Data
    {
        /// <summary>
        /// Header
        /// </summary>
        public Header header { get; set; }
        /// <summary>
        /// Content
        /// </summary>
        public Content content { get; set; }
    }

    public class Root
    {
        /// <summary>
        /// Code
        /// </summary>
        public string code { get; set; }
        /// <summary>
        /// success
        /// </summary>
        public string msg { get; set; }
        /// <summary>
        /// 40288a5f6518566601651861fd000004
        /// </summary>
        public string reqId { get; set; }
        /// <summary>
        /// Data
        /// </summary>
        public Data data { get; set; }
    }

    #endregion

    #region 在线查找

    public class MaGait
    {
        /// <summary>
        /// 1
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// ma10Parameters
        /// </summary>
        public string typeCode { get; set; }
        /// <summary>
        /// 起立过程时间
        /// </summary>
        public string typeName { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string leftNormal { get; set; }
        /// <summary>
        /// 22
        /// </summary>
        public string leftError { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string rightNormal { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string rightError { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string avgNormal { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string avgError { get; set; }
        /// <summary>
        /// cm
        /// </summary>
        public string unit { get; set; }

        //bangding
        public string reqId { get; set; }


    }

    public class Online_Data
    {
        /// <summary>
        /// 40288a5f64feff480164ff1f3606003c
        /// </summary>
        public string id { get; set; }
        /// <summary>
        /// 201808031729529029897
        /// </summary>
        public string serialNumber { get; set; }
        /// <summary>
        /// PatientId
        /// </summary>
        public string patientId { get; set; }
        /// <summary>
        /// 2018-08-03 17:19
        /// </summary>
        public string acquisitionStartTime { get; set; }
        /// <summary>
        /// 2018-08-03 17:19
        /// </summary>
        public string acquisitionEndTime { get; set; }
        /// <summary>
        /// EquipmentModel
        /// </summary>
        public string equipmentModel { get; set; }
        /// <summary>
        /// A22323
        /// </summary>
        public string equipmentNumber { get; set; }
        /// <summary>
        /// 40288a5f64feff480164ff14843a002e
        /// </summary>
        public string checkNumber { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string dataSources { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// 检查结果正常
        /// </summary>
        public string remark { get; set; }
        /// <summary>
        /// MaGait
        /// </summary>
        public List<MaGait> maGait { get; set; }

        //bangding
        public string reqId { get; set; }

    }

    public class Online_Root
    {
        /// <summary>
        /// Code
        /// </summary>
        public string code { get; set; }
        /// <summary>
        /// success
        /// </summary>
        public string msg { get; set; }
        /// <summary>
        /// 40288a5f6503b178016504465deb0245
        /// </summary>
        public string reqId { get; set; }
        /// <summary>
        /// Data
        /// </summary>
        public Online_Data data { get; set; }
    }

    public class OnlineShow
    {
        //////////////////// //Online_Root

        /// <summary>
        /// Code
        /// </summary>
        public string code { get; set; }
        /// <summary>
        /// success
        /// </summary>
        public string msg { get; set; }
        /// <summary>
        /// 40288a5f6503b178016504465deb0245
        /// </summary>
        public string reqId { get; set; }
        /// <summary>
        /// Data
        /// </summary>
        public Online_Data data { get; set; }


        ////////////////////  //Online_Data

        /// <summary>
        /// 40288a5f64feff480164ff1f3606003c
        /// </summary>
        public string id { get; set; }
        /// <summary>
        /// 201808031729529029897
        /// </summary>
        public string serialNumber { get; set; }
        /// <summary>
        /// PatientId
        /// </summary>
        public string patientId { get; set; }
        //new
        public string ptName { get; set; }

        /// <summary>
        /// 2018-08-03 17:19
        /// </summary>
        public string acquisitionStartTime { get; set; }
        /// <summary>
        /// 2018-08-03 17:19
        /// </summary>
        public string acquisitionEndTime { get; set; }
        /// <summary>
        /// EquipmentModel
        /// </summary>
        public string equipmentModel { get; set; }
        /// <summary>
        /// A22323
        /// </summary>
        public string equipmentNumber { get; set; }
        /// <summary>
        /// 40288a5f64feff480164ff14843a002e
        /// </summary>
        public string checkNumber { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string dataSources { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// 检查结果正常
        /// </summary>
        public string remark { get; set; }

        public string checkDoctor { get; set; }//-审核医生

        public string toExamineDoctor { get; set; }//-审核医生

        public string result { get; set; }//-审核医生

       
        ///////////////////////  //MaGait
        /// <summary>
        /// 1
        /// </summary>
        public string type_MaGait { get; set; }
        /// <summary>
        /// ma10Parameters
        /// </summary>
        public string typeCode { get; set; }
        /// <summary>
        /// 起立过程时间
        /// </summary>
        public string typeName { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string leftNormal { get; set; }
        /// <summary>
        /// 22
        /// </summary>
        public string leftError { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string rightNormal { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string rightError { get; set; }
        /// <summary>
        /// 2
        /// </summary>
        public string avgNormal { get; set; }
        /// <summary>
        /// 1
        /// </summary>
        public string avgError { get; set; }
        /// <summary>
        /// cm
        /// </summary>
        public string unit { get; set; }
    }
    public class OnlinecheckTableShow
    {
        public string name1 { get; set; }
        public string name2 { get; set; }
        public string name3 { get; set; }
        public string name4 { get; set; }
        public string name5{ get; set; }
        public string name6 { get; set; }
    
    }
    #endregion

    #region PDF

    public class ChildType
    {
        /// <summary>
        /// 1
        /// </summary>
        public string typeCode { get; set; }
        /// <summary>
        /// 起立过程时间
        /// </summary>
        public string typeName { get; set; }
        /// <summary>
        /// SiSt Duration
        /// </summary>
        public string typeEnName { get; set; }
        /// <summary>
        /// sec
        /// </summary>
        public string unit { get; set; }
        /// <summary>
        /// Distinguish
        /// </summary>
        public string distinguish { get; set; }

        public string partentID { get; set; }
    }

    public class Types
    {
        /// <summary>
        /// 1010
        /// </summary>
        public string typeCode { get; set; }
        /// <summary>
        /// 站立过程
        /// </summary>
        public string typeName { get; set; }
        /// <summary>
        /// SiSt
        /// </summary>
        public string typeEnName { get; set; }
        /// <summary>
        /// ChildType
        /// </summary>
        public List<ChildType> childType { get; set; }
    }

    public class PDF_Root
    {
        /// <summary>
        /// ma10Parameters
        /// </summary>
        public string typegroupcode { get; set; }
        /// <summary>
        /// Types
        /// </summary>
        public List<Types> types { get; set; }
    }




    #endregion
}
