using System;
using System.Collections.Generic;
using System.Linq;
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
        
    }
}
