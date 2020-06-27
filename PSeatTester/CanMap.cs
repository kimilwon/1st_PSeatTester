using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PSeatTester
{
    public class C_AVNIMSButtonCmd
    {
        public const int Addr = 0x4A1;
        public const int Lenfth = 3;
        public const int StartByte = 21;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            Memory_P1_CMD,
            /// <summary>
            /// 0x2
            /// </summary>
            Memory_P2_CMD,
            /// <summary>
            /// 0x3
            /// </summary>
            Memory_P3_CMD,
            /// <summary>
            /// 0x4
            /// </summary>
            PBack_P1_CMD,
            /// <summary>
            /// 0x5
            /// </summary>
            PBack_P2_CMD,
            /// <summary>
            /// 0x6
            /// </summary>
            PBack_P3_CMD,
            /// <summary>
            /// 0x7
            /// </summary>
            Invalid,
        }
    }

    public class LHD_Drv_SHVU_SeatHtOperSta
    {
        public const int Addr = 0x496;
        public const int Lenfth = 4;
        public const int StartByte = 3;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Heater_VentOff,
            /// <summary>
            /// 0x3
            /// </summary>
            VentLow,
            /// <summary>
            /// 0x4
            /// </summary>
            VentMid,
            /// <summary>
            /// 0x5
            /// </summary>
            VentHigh,
            /// <summary>
            /// 0x6
            /// </summary>
            HeaterLow,
            /// <summary>
            /// 0x7
            /// </summary>
            HeatherMid,
            /// <summary>
            /// 0x8
            /// </summary>
            HeatherHigh,
        }
    }

    public class RHD_Drv_SHVU_SeatHtOperSta
    {
        public const int Addr = 0x496;
        public const int Lenfth = 4;
        public const int StartByte = 35;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Heater_VentOff,
            /// <summary>
            /// 0x3
            /// </summary>
            VentLow,
            /// <summary>
            /// 0x4
            /// </summary>
            VentMid,
            /// <summary>
            /// 0x5
            /// </summary>
            VentHigh,
            /// <summary>
            /// 0x6
            /// </summary>
            HeaterLow,
            /// <summary>
            /// 0x7
            /// </summary>
            HeatherMid,
            /// <summary>
            /// 0x8
            /// </summary>
            HeatherHigh,
        }
    }

    public class LHD_Ast_SHVU_SeatHtOperSta
    {
        public const int Addr = 0x475;
        public const int Lenfth = 4;
        public const int StartByte = 3;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Heater_VentOff,
            /// <summary>
            /// 0x3
            /// </summary>
            VentLow,
            /// <summary>
            /// 0x4
            /// </summary>
            VentMid,
            /// <summary>
            /// 0x5
            /// </summary>
            VentHigh,
            /// <summary>
            /// 0x6
            /// </summary>
            HeaterLow,
            /// <summary>
            /// 0x7
            /// </summary>
            HeatherMid,
            /// <summary>
            /// 0x8
            /// </summary>
            HeatherHigh,
        }
    }

    public class RHD_Ast_SHVU_SeatHtOperSta
    {
        public const int Addr = 0x475;
        public const int Lenfth = 4;
        public const int StartByte = 35;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Heater_VentOff,
            /// <summary>
            /// 0x3
            /// </summary>
            VentLow,
            /// <summary>
            /// 0x4
            /// </summary>
            VentMid,
            /// <summary>
            /// 0x5
            /// </summary>
            VentHigh,
            /// <summary>
            /// 0x6
            /// </summary>
            HeaterLow,
            /// <summary>
            /// 0x7
            /// </summary>
            HeatherMid,
            /// <summary>
            /// 0x8
            /// </summary>
            HeatherHigh,
        }
    }

    public class Pass_SHVU_SeatHtOperSta
    {
        public const int Addr = 0x475;
        public const int Lenfth = 4;
        public const int StartByte = 3;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Heater_VentOff,
            /// <summary>
            /// 0x3
            /// </summary>
            VentLow,
            /// <summary>
            /// 0x4
            /// </summary>
            VentMid,
            /// <summary>
            /// 0x5
            /// </summary>
            VentHigh,
            /// <summary>
            /// 0x6
            /// </summary>
            HeaterLow,
            /// <summary>
            /// 0x7
            /// </summary>
            HeatherMid,
            /// <summary>
            /// 0x8
            /// </summary>
            HeatherHigh,
        }
    }

    public class RH__SHVU_SeatHtOperSta
    {

    }
    public class HU_FRSeatHeatVRcmd
    {
        public const int Addr = 0x4A1;
        public const int Lenfth = 4;
        public const int StartByte = 56;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Seat_Heat_Vent_OFF,
            /// <summary>
            /// 0x3
            /// </summary>
            Seat_Vent_Low,
            /// <summary>
            /// 0x4
            /// </summary>
            Seat_Vent_Mid,
            /// <summary>
            /// 0x5
            /// </summary>
            Seat_Vent_High,
            /// <summary>
            /// 0x6
            /// </summary>
            Seat_Heat_Low,
            /// <summary>
            /// 0x7
            /// </summary>
            Seat_Heat_Mid,
            /// <summary>
            /// 0x8
            /// </summary>
            Seat_Heat_High,
            /// <summary>
            /// 0x9~0xE
            /// </summary>
            Unused2,
            /// <summary>
            /// 0xF
            /// </summary>
            Invalid = 15
        }
    }
    public class HU_FLSeatHeatVRcmd
    {
        public const int Addr = 0x4A1;
        public const int Lenfth = 4;
        public const int StartByte = 60;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Init,
            /// <summary>
            /// 0x1
            /// </summary>
            Unused,
            /// <summary>
            /// 0x2
            /// </summary>
            Seat_Heat_Vent_OFF,
            /// <summary>
            /// 0x3
            /// </summary>
            Seat_Vent_Low,
            /// <summary>
            /// 0x4
            /// </summary>
            Seat_Vent_Mid,
            /// <summary>
            /// 0x5
            /// </summary>
            Seat_Vent_High,
            /// <summary>
            /// 0x6
            /// </summary>
            Seat_Heat_Low,
            /// <summary>
            /// 0x7
            /// </summary>
            Seat_Heat_Mid,
            /// <summary>
            /// 0x8
            /// </summary>
            Seat_Heat_High,
            /// <summary>
            /// 0x9~0xE
            /// </summary>
            Unused2,
            /// <summary>
            /// 0xF
            /// </summary>
            Invalid = 15
        }
    }
    public class C_FRSeatCoolerSW
    {
        public const int Addr = 0x4CA;
        public const int Lenfth = 2;
        public const int StartByte = 8;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Reserved,
            /// <summary>
            /// 0x3
            /// </summary>
            invalid
        }
    }
    public class C_FRSeatHeaterSW
    {
        public const int Addr = 0x4CA;
        public const int Lenfth = 2;
        public const int StartByte = 10;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Reserved,
            /// <summary>
            /// 0x3
            /// </summary>
            invalid
        }
    }
    public class C_FLSeatCoolerSW
    {
        public const int Addr = 0x4CA;
        public const int Lenfth = 2;
        public const int StartByte = 12;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Reserved,
            /// <summary>
            /// 0x3
            /// </summary>
            invalid
        }
    }
    public class C_FLSeatHeaterSW
    {
        public const int Addr = 0x4CA;
        public const int Lenfth = 2;
        public const int StartByte = 14;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Reserved,
            /// <summary>
            /// 0x3
            /// </summary>
            invalid
        }
    }
    public class IMS_DrvrImsSwSetSta
    {
        public const int Addr = 0x402;
        public const int Lenfth = 2;
        public const int StartByte = 16;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class IMS_DrvrImsSw1Sta
    {
        public const int Addr = 0x402;
        public const int Lenfth = 2;
        public const int StartByte = 18;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class IMS_DrvrImsSw2Sta
    {
        public const int Addr = 0x402;
        public const int Lenfth = 2;
        public const int StartByte = 20;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }

    public class DrvIMS_MemoryP1Req
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 0;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1
            /// </summary>
            P1,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class DrvIMS_MemoryP2Req
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 2;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1
            /// </summary>
            P2,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class DrvIMS_MemorySetEnaSta
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 6;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Disable,
            /// <summary>
            /// 0x1
            /// </summary>
            Enable,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class DrvIMS_PlyBckP1Req
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 8;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1
            /// </summary>
            P1,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class DrvIMS_PlyBckP2Req
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 10;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1
            /// </summary>
            P3,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }

    public class DrvIMS_PlyBckStpReq
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 14;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1
            /// </summary>
            Stop,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator

        }
    }
    public class SmrtIMS_PlyBckReq
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 12;
        public const int StartByte = 16;

        public enum Data : ushort
        {
            /// <summary>
            /// 0x0
            /// </summary>
            None,
            /// <summary>
            /// 0x1 ~ 0x97E (1 ~ 2430) 까지 스마트 IMS DB 번호
            /// </summary>
            DbNo,
            Invalid = 0xfff
        }
    }

    public class SmrtIMS_PlyBckStpReq
    {
        public const int Addr = 0x4cb;
        public const int Lenfth = 2;
        public const int StartByte = 37;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            On,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class PSM_DrvStEsyAcsUSMSta
    {
        public const int Addr = 0x386;
        public const int Lenfth = 3;
        public const int StartByte = 0;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0
            /// </summary>
            Default_Value,
            //0x1 (Disable)
            Off,
            //0x2 (Enable)
            Storke_50mmOn,
            /// <summary>
            /// 0x3 (Enable)
            /// </summary>
            Storke_75mmOn,
            /// <summary>
            /// 0x4
            /// </summary>
            Reserved1,
            /// <summary>
            /// 0x5
            /// </summary>
            Reserved2,
            /// <summary>
            /// 0x6
            /// </summary>
            Reserved3,
            /// <summary>
            /// 0x7
            /// </summary>
            Invalid
        }
    }
    public class PSM_DrvStSldFwdDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 0;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStSldBckwdDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 2;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStReclneFwdDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 4;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStReclneBckwdDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 6;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStTiltUpDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 8;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStTiltDnDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 10;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStHghtUpDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 12;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStHghtDnDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 14;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }

    public class PSM_DrvStLmbrDefDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 32;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStLmbrLoDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 34;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStLmbrMdlDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 36;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStLmbrUpDis
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 38;

        public enum Data : byte
        {
            ///0x0
            Default,
            /// <summary>
            /// 0x1
            /// </summary>
            off,
            /// <summary>
            /// 0x2
            /// </summary>
            on,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }


    public class PSM_LmbrSigSrcTyp
    {
        public const int Addr = 0x387;
        public const int Lenfth = 2;
        public const int StartByte = 44;

        public enum Data : byte
        {
            ///0x0
            Not_PSM,
            /// <summary>
            /// 0x1
            /// </summary>
            PSM,
            /// <summary>
            /// 0x2
            /// </summary>
            reserved,
            /// <summary>
            /// 0x3
            /// </summary>
            nvalid
        }
    }
    public class PSM_DrvStReclnePosVal
    {
        public const int Addr = 0x388;
        public const int Lenfth = 3;
        public const int StartByte = 0;

        public enum Data : byte
        {
            /// <summary>
            /// 0x0 - 0≤x≤33
            /// </summary>
            Value1,
            /// <summary>
            /// 0x1 - 33<x≤41
            /// </summary>
            Value2,
            /// <summary>
            /// 0x2 - 41<x≤46
            /// </summary>
            Value3,
            /// <summary>
            /// 0x3 - 46<x≤53
            /// </summary>
            Value4,
            /// <summary>
            /// 0x4 - 53<x≤66
            /// </summary>
            Value5,
            /// <summary>
            /// 0x5
            /// </summary>
            reserved1,
            /// <summary>
            /// 0x6
            /// </summary>
            reserved2,
            /// <summary>
            /// 0x7
            /// </summary>
            invalid
        }

        public string DataValue(Data Value)
        {
            string[] DataToString = { "0≤x≤33", "33<x≤41", "41<x≤46", "46<x≤53", "53 < x≤66" };

            if ((int)Value < DataToString.Length)
                return DataToString[(int)Value];
            else return "reserved";
        }
    }
    public class PSM_DrvStSldPosVal
    {
        public const int Addr = 0x388;
        public const int Lenfth = 3;
        public const int StartByte = 3;


        public enum Data : byte
        {
            /// <summary>
            /// 0x0 - 0≤x≤130
            /// </summary>
            Value1,
            /// <summary>
            /// 0x1 - 130<x≤180
            /// </summary>
            Value2,
            /// <summary>
            /// 0x2 - 180<x≤230
            /// </summary>
            Value3,
            /// <summary>
            /// 0x3 - 230<x≤260
            /// </summary>
            Value4,
            /// <summary>
            /// 0x4 - 53<x≤66
            /// </summary>
            reserved1,
            /// <summary>
            /// 0x5
            /// </summary>
            reserved2,
            /// <summary>
            /// 0x6
            /// </summary>
            reserved3,
            /// <summary>
            /// 0x7
            /// </summary>
            invalid
        }

        static public string DataValue(Data Value)
        {
            string[] DataToString = { "0≤x≤130", "130<x≤180", "180<x≤230", "230<x≤260" };

            if ((int)Value < DataToString.Length)
                return DataToString[(int)Value];
            else return "reserved";
        }
    }
    public class PSM_DrvStMvSta
    {
        public const int Addr = 0x388;
        public const int Lenfth = 2;
        public const int StartByte = 6;

        public enum Data : byte
        {
            ///0x0
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class PSM_EsyAcsGetOnSta
    {
        public const int Addr = 0x388;
        public const int Lenfth = 2;
        public const int StartByte = 8;

        public enum Data : byte
        {
            ///0x0
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class PSM_EsyAcsGetOffSta
    {
        public const int Addr = 0x388;
        public const int Lenfth = 2;
        public const int StartByte = 10;

        public enum Data : byte
        {
            ///0x0
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class PSM_EsyAcsStpSta
    {
        public const int Addr = 0x388;
        public const int Lenfth = 2;
        public const int StartByte = 12;

        public enum Data : byte
        {
            ///0x0
            Off,
            /// <summary>
            /// 0x1
            /// </summary>
            on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class BCM_AccInSta
    {
        public const int Addr = 0x3E0;
        public const int Lenfth = 2;
        public const int StartByte = 34;

        public enum Data : byte
        {
            ///0x0
            Acc_Off,
            /// <summary>
            /// 0x1
            /// </summary>
            Acc_on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class BCM_Ign1InSta
    {
        public const int Addr = 0x3E0;
        public const int Lenfth = 2;
        public const int StartByte = 36;

        public enum Data : byte
        {
            ///0x0
            IGN1_Off,
            /// <summary>
            /// 0x1
            /// </summary>
            IGN1_on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }
    public class BCM_Ign2InSta
    {
        public const int Addr = 0x3E0;
        public const int Lenfth = 2;
        public const int StartByte = 38;

        public enum Data : byte
        {
            ///0x0
            IGN2_Off,
            /// <summary>
            /// 0x1
            /// </summary>
            IGN2_on,
            /// <summary>
            /// 0x2
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x3
            /// </summary>
            Error_Indicator
        }
    }

    public class BCM_IgnSwSta
    {
        public const int Addr = 0x3E0;
        public const int Lenfth = 3;
        public const int StartByte = 40;

        public enum Data : byte
        {
            ///0x0
            KeyOff,
            /// <summary>
            /// 0x1
            /// </summary>
            KeyIn,
            /// <summary>
            /// 0x2
            /// </summary>
            Acc,
            /// <summary>
            /// 0x3
            /// </summary>
            Ign,
            /// <summary>
            /// 0x4
            /// </summary>
            St,
            /// <summary>
            /// 0x5
            /// </summary>
            Reseved,
            /// <summary>
            /// 0x6
            /// </summary>
            Not_Used,
            /// <summary>
            /// 0x7
            /// </summary>
            ErrorIndicator
        }
    }

    public class PSeat_Slide_Fwd
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x11, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x11, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x11, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Slide_Bwd
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x12, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x12, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Recline_Fwd
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x13, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x13, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x13, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Recline_Bwd
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x14, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x14, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Tilt_Up
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x15, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x15, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x15, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Tilt_Down
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x16, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x16, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x16, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Height_Up
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x17, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x17, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x17, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class PSeat_Height_Down
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] MoveData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x18, 0x03, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] StopData
        {
            get
            {
                byte[] Data = { 0x04, 0x2f, 0xf0, 0x18, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
        static public byte[] ResponseData
        {
            get
            {
                byte[] Data = { 0x6f, 0xf0, 0x18, 0x00, 0x00, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }

    public class VirtialLimite_All
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] Clear
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x02, 0x12, 0xa3, 0x80, 0x00, 0x00 };

                return Data;
            }
        }

        static public byte[] LimitSet
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x01, 0x12, 0xa3, 0x00, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class VirtialLimite_Slide
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;
        static public byte[] Clear
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x02, 0x12, 0xa3, 0x81, 0x00, 0x00 };

                return Data;
            }
        }

        static public byte[] LimitSet
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x01, 0x12, 0xa3, 0x01, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class VirtialLimite_Recline
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] Clear
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x02, 0x12, 0xa3, 0x82, 0x00, 0x00 };

                return Data;
            }
        }

        static public byte[] LimitSet
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x01, 0x12, 0xa3, 0x02, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class VirtialLimite_Tilt
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;
        static public byte[] Clear
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x02, 0x12, 0xa3, 0x83, 0x00, 0x00 };

                return Data;
            }
        }

        static public byte[] LimitSet
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x01, 0x12, 0xa3, 0x03, 0x00, 0x00 };

                return Data;
            }
        }
    }
    public class VirtialLimite_Height
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;
        static public byte[] Clear
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x02, 0x12, 0xa3, 0x84, 0x00, 0x00 };

                return Data;
            }
        }

        static public byte[] LimitSet
        {
            get
            {
                byte[] Data = { 0x05, 0x31, 0x01, 0x12, 0xa3, 0x04, 0x00, 0x00 };

                return Data;
            }
        }
    }

    public class VirtialLimite_GetMessage
    {
        //public const int RequestAddr = 0x726;
        //public const int ResponseAddr = 0x706;
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] GetRequest
        {
            get
            {
                //byte[] Data = { 0x05, 0x31, 0x03, 0x12, 0xa3, 0x0, 0x00, 0x00};
                byte[] Data = { 0x04, 0x31, 0x03, 0x12, 0xa3, 0x0, 0x00, 0x00 };

                return Data;
            }
        }
    }
    /// <summary>
    /// LH 소프트웨어 버전 입력
    /// </summary>
    public class SWVersionRead
    {
        public const int RequestAddr = 0x7a3;
        public const int ResponseAddr = 0x7ab;

        static public byte[] GetRequest
        {
            get
            {
                byte[] Data = { 0x03, 0x22, 0xf1, 0xa0, 0xaa, 0xaa, 0xaa, 0xaa };

                return Data;
            }
        }
    }

    public enum IN_CAN_LIST
    {
        LHD_Drv_SHVU_SeatHtOperSta,
        RHD_Drv_SHVU_SeatHtOperSta,
        LHD_Ast_SHVU_SeatHtOperSta,
        RHD_Ast_SHVU_SeatHtOperSta,
        Drv_SHVU_SeatHtOperSta,
        RH__SHVU_SeatHtOperSta,
        C_AVNIMSButtonCmd,
        HU_FRSeatHeatVRcmd,
        HU_FLSeatHeatVRcmd,
        C_FRSeatCoolerSW,
        C_FRSeatHeaterSW,
        C_FLSeatCoolerSW,
        C_FLSeatHeaterSW,
        IMS_DrvrImsSwSetSta,
        IMS_DrvrImsSw1Sta,
        IMS_DrvrImsSw2Sta,
        DrvIMS_MemoryP1Req,
        DrvIMS_MemoryP2Req,
        DrvIMS_MemorySetEnaSta,
        DrvIMS_PlyBckP1Req,
        DrvIMS_PlyBckP2Req,
        DrvIMS_PlyBckStpReq,
        SmrtIMS_PlyBckReq,
        SmrtIMS_PlyBckStpReq,
        PSM_DrvStEsyAcsUSMSta,
        PSM_DrvStSldFwdDis,
        PSM_DrvStSldBckwdDis,
        PSM_DrvStReclneFwdDis,
        PSM_DrvStReclneBckwdDis,
        PSM_DrvStTiltUpDis,
        PSM_DrvStTiltDnDis,
        PSM_DrvStHghtUpDis,
        PSM_DrvStHghtDnDis,
        PSM_DrvStLmbrDefDis,
        PSM_DrvStLmbrLoDis,
        PSM_DrvStLmbrMdlDis,
        PSM_DrvStLmbrUpDis,
        PSM_LmbrSigSrcTyp,
        PSM_DrvStReclnePosVal,
        PSM_DrvStSldPosVal,
        PSM_DrvStMvSta,
        PSM_EsyAcsGetOnSta,
        PSM_EsyAcsGetOffSta,
        PSM_EsyAcsStpSta,
        BCM_AccInSta,
        BCM_Ign1InSta,
        BCM_Ign2InSta,
        BCM_IgnSwSta,
        PSeat_Slide_Fwd,
        PSeat_Slide_Bwd,
        PSeat_Recline_Fwd,
        PSeat_Recline_Bwd,
        PSeat_Tilt_Up,
        PSeat_Tilt_Down,
        PSeat_Height_Up,
        PSeat_Height_Down,
        VirtialLimite_Slide,
        VirtialLimite_Recline,
        VirtialLimite_Tilt,
        VirtialLimite_Height
    }
    public enum OUT_CAN_LIST
    {
        C_AVNIMSButtonCmd,
        HU_FRSeatHeatVRcmd,
        HU_FLSeatHeatVRcmd,
        C_FRSeatCoolerSW,
        C_FRSeatHeaterSW,
        C_FLSeatCoolerSW,
        C_FLSeatHeaterSW,
        IMS_DrvrImsSwSetSta,
        IMS_DrvrImsSw1Sta,
        IMS_DrvrImsSw2Sta,
        DrvIMS_MemoryP1Req,
        DrvIMS_MemoryP2Req,
        DrvIMS_MemorySetEnaSta,
        DrvIMS_PlyBckP1Req,
        DrvIMS_PlyBckP2Req,
        DrvIMS_PlyBckStpReq,
        SmrtIMS_PlyBckReq,
        SmrtIMS_PlyBckStpReq,
        PSM_DrvStEsyAcsUSMSta,
        PSM_DrvStSldFwdDis,
        PSM_DrvStSldBckwdDis,
        PSM_DrvStReclneFwdDis,
        PSM_DrvStReclneBckwdDis,
        PSM_DrvStTiltUpDis,
        PSM_DrvStTiltDnDis,
        PSM_DrvStHghtUpDis,
        PSM_DrvStHghtDnDis,
        PSM_DrvStLmbrDefDis,
        PSM_DrvStLmbrLoDis,
        PSM_DrvStLmbrMdlDis,
        PSM_DrvStLmbrUpDis,
        PSM_LmbrSigSrcTyp,
        PSM_DrvStReclnePosVal,
        PSM_DrvStSldPosVal,
        PSM_DrvStMvSta,
        PSM_EsyAcsGetOnSta,
        PSM_EsyAcsGetOffSta,
        PSM_EsyAcsStpSta,
        BCM_AccInSta,
        BCM_Ign1InSta,
        BCM_Ign2InSta,
        BCM_IgnSwSta,
        PSeat_Slide_Fwd,
        PSeat_Slide_Bwd,
        PSeat_Recline_Fwd,
        PSeat_Recline_Bwd,
        PSeat_Tilt_Up,
        PSeat_Tilt_Down,
        PSeat_Height_Up,
        PSeat_Height_Down,
        VirtialLimite_All,
        VirtialLimite_Slide,
        VirtialLimite_Recline,
        VirtialLimite_Tilt,
        VirtialLimite_Height,
        VirtialLimite_GetMessage,
    }

    public enum IN_LIN_LIST
    {
        SWEvenLEDDimReq,
        RH_VentSWLEDMid,
        LH_VentSWLEDHigh,
        LH_VentSWLEDLow,
        RH_VentSWLEDHigh,
        RH_VentSWLEDLow,
        LH_VentSWLEDMid,
        SWOddLEDDimReq,
        LH_HeaterSWLEDMid,
        RH_HeaterSWLEDHigh,
        RH_HeaterSWLEDLow,
        LH_HeaterSWLEDLow,
        LH_HeaterSWLEDHigh,
        RH_HeaterSWLEDMid
    }
    public enum OUT_LIN_LIST
    {
        LH_HeaterSW,
        RH_HeaterSW,
        LH_VentSW,
        RH_VentSW,
        LH_HeaterSWRaw,
        RH_HeaterSWRaw,
        SWEvenLEDDimRes,
        SWLINError,
        LH_VentSWRaw,
        RH_VentSWRaw,
        SWOddLEDDimRes,
        SWLEDStatus
    }
    /// <summary>
    /// 스위치 모듈의 짝수 번째 LED Port(LED0, LED2,LED4,LED6,LED8, LED10)의 Dimming 제어Request 정보
    /// </summary>
    public class SWEvenLEDDimReq
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 7;

        public enum Data : byte
        {
            /// <summary>
            /// Initila Value
            /// </summary>
            DimmingControlByPWM,
            /// <summary>
            /// 100 % Dimming
            /// </summary>
            Fixed100ProDimming
        }
    }
    /// <summary>
    /// RH 통풍 스위치 Mid Indicator(LED4) 정보
    /// </summary>
    public class RH_VentSWLEDMid
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 6;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }/// <summary>
     /// LH 통풍 스위치 High Indicator(LED2) 정보
     /// </summary>
    public class LH_VentSWLEDHigh
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 5;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// LH 통풍 스위치 Low Indicator(LED0) 정보
    /// </summary>
    public class LH_VentSWLEDLow
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 4;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }/// <summary>
     /// RH 통풍 스위치 High Indicator(LED5) 정보
     /// </summary>
    public class RH_VentSWLEDHigh
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 2;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }/// <summary>
     /// RH 통풍 스위치 Low Indicator(LED3) 정보
     /// </summary>
    public class RH_VentSWLEDLow
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 1;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// LH 통풍 스위치 Mid Indicator(LED1) 정보
    /// </summary>
    public class LH_VentSWLEDMid
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 0;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// 스위치 모듈의홀수 번째 LED Port(LED1, LED3,LED5,LED7,LED9, LED11)의 Dimming 제어 Request 정보
    /// 0x00 : Dimming Control by PWM(Initial Value)						
    /// 0x01 : Fixed 100% Dimming
    /// </summary>
    public class SWOddLEDDimReq
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 15;

        public enum Data : byte
        {
            /// <summary>
            /// Initila Value
            /// </summary>
            DimmingControlByPWM,
            /// <summary>
            /// 100 % Dimming
            /// </summary>
            Fixed100ProDimming
        }
    }
    /// <summary>
    /// LH 히터 스위치 Mid Indicator(LED10) 정보
    /// </summary>
    public class LH_HeaterSWLEDMid
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 14;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// RH 히터 스위치 High Indicator(LED8) 정보
    /// </summary>
    public class RH_HeaterSWLEDHigh
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 13;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }/// <summary>
     /// RH 히터 스위치 Low Indicator(LED6) 정보
     /// </summary>
    public class RH_HeaterSWLEDLow
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 12;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// LH 히터 스위치 Low Indicator(LED11) 정보
    /// </summary>
    public class LH_HeaterSWLEDLow
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 10;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// LH 히터 스위치 High Indicator(LED9) 정보
    /// </summary>
    public class LH_HeaterSWLEDHigh
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 9;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }
    /// <summary>
    /// RH 히터 스위치 Mid Indicator(LED7) 정보
    /// </summary>
    public class RH_HeaterSWLEDMid
    {
        public const int Addr = 0x08;
        public const int Lenfth = 1;
        public const int StartByte = 8;

        public enum Data : byte
        {
            /// <summary>
            /// LED Off
            /// </summary>
            Off,
            /// <summary>
            /// LED On
            /// </summary>
            On
        }
    }


    /// <summary>
    /// 스위치 모듈에서 LH 히터스위치 입력 신호의 채터링현상을 처리한 이후의 정보
    /// </summary>
    public class LH_HeaterSW
    {
        public const int Addr = 0x04;
        public const int Lenfth = 2;
        public const int StartByte = 6;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : Initial Value
            /// </summary>
            No_Action,
            /// <summary>
            /// 0x01 :STP
            /// </summary>
            Short_Term_Push,
            /// <summary>
            /// 0x02 : LTP
            /// </summary>
            Long_Term_Push,
            /// <summary>
            /// 0x03 : DP
            /// </summary>
            Diagnostic_Push
        }
    }
    /// <summary>
    /// 스위치 모듈에서 RH 히터스위치 입력 신호의 채터링현상을 처리한 이후의 정보
    /// </summary>
    public class RH_HeaterSW
    {
        public const int Addr = 0x04;
        public const int Lenfth = 2;
        public const int StartByte = 4;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : Initial Value
            /// </summary>
            No_Action,
            /// <summary>
            /// 0x01 :STP
            /// </summary>
            Short_Term_Push,
            /// <summary>
            /// 0x02 : LTP
            /// </summary>
            Long_Term_Push,
            /// <summary>
            /// 0x03 : DP
            /// </summary>
            Diagnostic_Push
        }
    }
    /// <summary>
    /// 스위치 모듈에서 LH 통풍스위치 입력 신호의 채터링현상을 처리한 이후의 정보
    /// </summary>
    public class LH_VentSW
    {
        public const int Addr = 0x04;
        public const int Lenfth = 2;
        public const int StartByte = 2;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : Initial Value
            /// </summary>
            No_Action,
            /// <summary>
            /// 0x01 :STP
            /// </summary>
            Short_Term_Push,
            /// <summary>
            /// 0x02 : LTP
            /// </summary>
            Long_Term_Push,
            /// <summary>
            /// 0x03 : DP
            /// </summary>
            Diagnostic_Push
        }
    }
    /// <summary>
    /// 스위치 모듈에서 RH 통풍스위치 입력 신호의 채터링현상을 처리한 이후의 정보
    /// </summary>
    public class RH_VentSW
    {
        public const int Addr = 0x04;
        public const int Lenfth = 2;
        public const int StartByte = 0;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : Initial Value
            /// </summary>
            No_Action,
            /// <summary>
            /// 0x01 :STP
            /// </summary>
            Short_Term_Push,
            /// <summary>
            /// 0x02 : LTP
            /// </summary>
            Long_Term_Push,
            /// <summary>
            /// 0x03 : DP
            /// </summary>
            Diagnostic_Push
        }
    }
    /// <summary>
    /// 스위치 모듈에서 LH 히터스위치 입력 신호의 채터링 현상을 처리하기 이전의 Raw Data
    /// </summary>
    public class LH_HeaterSWRaw
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 15;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : On(Initial Value) 
            /// </summary>
            On,
            /// <summary>
            /// 0x01 : Off
            /// </summary>
            Off
        }
    }
    /// <summary>
    /// 스위치 모듈에서 RH 히터스위치 입력 신호의 채터링 현상을 처리하기 이전의 Raw Data
    /// </summary>
    public class RH_HeaterSWRaw
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 14;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : On(Initial Value) 
            /// </summary>
            On,
            /// <summary>
            /// 0x01 : Off
            /// </summary>
            Off
        }
    }
    /// <summary>
    /// 스위치 모듈의 짝수 번째 LED Port(LED0, LED2, LED4, LED6, LED8, LED10)의 Dimming 제어Response 정보
    /// </summary>
    public class SWEvenLEDDimRes
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 13;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00: Dimming Control by PWM(Initial Value)
            /// </summary>
            DimmingControlByPWM,
            /// <summary>
            ///  0x01: Fixed 100% Dimming 
            /// </summary>
            Fixed_100Pro_Dimming
        }
    }
    /// <summary>
    /// 스위치 LIN통신 Bit Error 정보
    /// </summary>
    public class SWLINError
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 12;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00: Dimming Control by PWM(Initial Value)
            /// </summary>
            Normal,
            /// <summary>
            ///  0x01: Fixed 100% Dimming 
            /// </summary>
            BitError
        }
    }
    /// <summary>
    /// 스위치 모듈에서 LH 통풍스위치 입력 신호의 채터링 현상을 처리하기 이전의 Raw Data
    /// </summary>
    public class LH_VentSWRaw
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 11;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : On(Initial Value) 
            /// </summary>
            On,
            /// <summary>
            /// 0x01 : Off
            /// </summary>
            Off
        }
    }
    /// <summary>
    /// 스위치 모듈에서 RH 통풍스위치 입력 신호의 채터링 현상을 처리하기 이전의 Raw Data
    /// </summary>
    public class RH_VentSWRaw
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 10;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00 : On(Initial Value) 
            /// </summary>
            On,
            /// <summary>
            /// 0x01 : Off
            /// </summary>
            Off
        }
    }
    /// <summary>
    /// 스위치 모듈의 홀수 번째 LED Port(LED1, LED3,LED5, LED7, LED9, LED11)의 Dimming 제어Response 정보
    /// </summary>
    public class SWOddLEDDimRes
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 9;

        public enum Data : byte
        {
            /// <summary>
            /// 0x00: Dimming Control by PWM(Initial Value)
            /// </summary>
            DimmingControlByPWM,
            /// <summary>
            ///  0x01: Fixed 100% Dimming 
            /// </summary>
            Fixed_100Pro_Dimming
        }
    }
    /// <summary>
    /// 스위치 LED Status
    /// </summary>
    public class SWLEDStatus
    {
        public const int Addr = 0x04;
        public const int Lenfth = 1;
        public const int StartByte = 8;

        public enum Data : byte
        {
            /// <summary>
            LEDChannelTurnOnStatus,
            AllLEDChannelTurnOffStatus
        }
    }



    public class CanMap
    {
        private __CanLin__ mCan;
        private MyInterface mControl = null;



        public __CanMsg PositionRead = new __CanMsg()
        {
            DATA = new byte[8]
        };

        public __CanMsg PSeatSwRead = new __CanMsg()
        {
            DATA = new byte[8]
        };
        public __CanMsg MotorStatusRead = new __CanMsg()
        {
            DATA = new byte[8]
        };

        public __CanMsg PSeatControlStart = new __CanMsg()
        {
            DATA = new byte[8]
        };

        public __CanMsg VirtialLimite_CommunicationStart = new __CanMsg()
        {
            DATA = new byte[8]
        };

        //public __CanMsg VirtualLimetRequest = new __CanMsg()
        //{
        //    DATA = new byte[8]
        //};

        public CanMap()
        {
            ThreadSetting();
        }
        public CanMap(MyInterface mControl)
        {
            this.mControl = mControl;
            ThreadSetting();
            mCan.Can.In.Can.Send = new __SendCan__[20];
            mCan.Can.Out.Can.Send = new __SendCan__[20];

            mCan.Lin.In.Lin.Send = new __SendCan__[20];
            mCan.Lin.Out.Lin.Send = new __SendCan__[20];

            for (int i = 0; i < 20; i++)
            {
                mCan.Can.In.Can.Send[i].Data = new byte[8];
                mCan.Can.Out.Can.Send[i].Data = new byte[8];

                mCan.Lin.In.Lin.Send[i].Data = new byte[8];
                mCan.Lin.Out.Lin.Send[i].Data = new byte[8];
            }
        }
        ~CanMap()
        {

        }

        public void CanLinDefaultSetting()
        {
            mCan.Can.In.Can.Send[0].ID = 0x506;
            mCan.Can.In.Can.Send[0].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[0].Data, 0, 8);

            mCan.Can.In.Can.Send[1].ID = 0x386;
            mCan.Can.In.Can.Send[1].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[1].Data, 0, 8);

            mCan.Can.In.Can.Send[2].ID = 0x387;
            mCan.Can.In.Can.Send[2].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[2].Data, 0, 8);

            mCan.Can.In.Can.Send[3].ID = 0x388;
            mCan.Can.In.Can.Send[3].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[3].Data, 0, 8);

            mCan.Can.In.Can.Send[4].ID = 0x4cb;
            mCan.Can.In.Can.Send[4].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[4].Data, 0, 8);

            mCan.Can.In.Can.Send[5].ID = 0x7ab;
            mCan.Can.In.Can.Send[5].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[5].Data, 0, 8);

            mCan.Can.In.Can.Send[6].ID = 0x71;
            mCan.Can.In.Can.Send[6].Length = 8;
            Array.Clear(mCan.Can.In.Can.Send[6].Data, 0, 8);

            mCan.Can.In.Can.Max = 7;

            mCan.Can.Out.Can.Send[0].ID = 0x7a3;
            mCan.Can.Out.Can.Send[0].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[0].Data, 0, 8);

            mCan.Can.Out.Can.Send[1].ID = 0x3e0;
            mCan.Can.Out.Can.Send[1].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[1].Data, 0, 8);

            mCan.Can.Out.Can.Send[2].ID = 0x444;
            mCan.Can.Out.Can.Send[2].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[2].Data, 0, 8);

            mCan.Can.Out.Can.Send[3].ID = 0x31;
            mCan.Can.Out.Can.Send[3].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[3].Data, 0, 8);

            mCan.Can.Out.Can.Send[4].ID = 0x402;
            mCan.Can.Out.Can.Send[4].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[4].Data, 0, 8);

            mCan.Can.Out.Can.Send[5].ID = 0x4A1;
            mCan.Can.Out.Can.Send[5].Length = 8;
            Array.Clear(mCan.Can.Out.Can.Send[5].Data, 0, 8);

            mCan.Can.Out.Can.Max = 6;

            PositionRead.ID = 0x7a3;
            PositionRead.Length = 8;

            PositionRead.DATA[0] = 0x03;
            PositionRead.DATA[1] = 0x22;
            PositionRead.DATA[2] = 0xb4;
            PositionRead.DATA[3] = 0x02;
            PositionRead.DATA[4] = 0x00;
            PositionRead.DATA[5] = 0x00;
            PositionRead.DATA[6] = 0x00;
            PositionRead.DATA[7] = 0x00;

            PSeatSwRead.ID = 0x7a3;
            PSeatSwRead.Length = 8;

            PSeatSwRead.DATA[0] = 0x03;
            PSeatSwRead.DATA[1] = 0x22;
            PSeatSwRead.DATA[2] = 0xb4;
            PSeatSwRead.DATA[3] = 0x03;
            PSeatSwRead.DATA[4] = 0x00;
            PSeatSwRead.DATA[5] = 0x00;
            PSeatSwRead.DATA[6] = 0x00;
            PSeatSwRead.DATA[7] = 0x00;

            MotorStatusRead.ID = 0x7a3;
            MotorStatusRead.Length = 8;


            MotorStatusRead.DATA[0] = 0x03;
            MotorStatusRead.DATA[1] = 0x22;
            MotorStatusRead.DATA[2] = 0xb4;
            MotorStatusRead.DATA[3] = 0x04;
            MotorStatusRead.DATA[4] = 0x00;
            MotorStatusRead.DATA[5] = 0x00;
            MotorStatusRead.DATA[6] = 0x00;
            MotorStatusRead.DATA[7] = 0x00;

            PSeatControlStart.ID = 0x7a3;
            PSeatControlStart.Length = 8;
            PSeatControlStart.DATA[0] = 0x02;
            PSeatControlStart.DATA[1] = 0x10;
            PSeatControlStart.DATA[2] = 0x03;
            PSeatControlStart.DATA[4] = 0x00;
            PSeatControlStart.DATA[4] = 0x00;
            PSeatControlStart.DATA[5] = 0x00;
            PSeatControlStart.DATA[6] = 0x00;
            PSeatControlStart.DATA[7] = 0x00;

            //가상 리미트 여부를 가져온다.
            VirtialLimite_CommunicationStart.ID = 0x7a3;
            VirtialLimite_CommunicationStart.Length = 8;
            VirtialLimite_CommunicationStart.DATA[0] = 0x03;
            VirtialLimite_CommunicationStart.DATA[1] = 0x19;
            VirtialLimite_CommunicationStart.DATA[2] = 0x02;
            VirtialLimite_CommunicationStart.DATA[3] = 0x08;
            VirtialLimite_CommunicationStart.DATA[4] = 0x00;
            VirtialLimite_CommunicationStart.DATA[5] = 0x00;
            VirtialLimite_CommunicationStart.DATA[6] = 0x00;
            VirtialLimite_CommunicationStart.DATA[7] = 0x00;

            Array.Clear(MultiFreameData, 0, MultiFreameData.Length);
            MultiFrameLength = 0;
            MultiFrameRead = false;
            MultiFreameDataPos = 0;

            //VirtualLimetRequest.ID = 0x7a3;
            //VirtualLimetRequest.Length = 8;

            //VirtualLimetRequest.DATA[0] = 0x04;
            //VirtualLimetRequest.DATA[1] = 0x31;
            //VirtualLimetRequest.DATA[2] = 0x03;
            //VirtualLimetRequest.DATA[3] = 0x12;
            //VirtualLimetRequest.DATA[4] = 0xa3;
            //VirtualLimetRequest.DATA[5] = 0x00;
            //VirtualLimetRequest.DATA[6] = 0x00;
            //VirtualLimetRequest.DATA[7] = 0x00;

            SwVerData = "";
            CanFailFlag = false;
            CanFailCheckFirstTime = mControl.공용함수.timeGetTimems();
            CanFailCheckLastTime = mControl.공용함수.timeGetTimems();
            return;
        }

        private bool CanFailFlag = false;
        private long CanFailCheckFirstTime = 0;
        private long CanFailCheckLastTime = 0;

        public byte CheckCanInMessage(IN_CAN_LIST Msg)
        {
            __CanInPos Pos;// = new __LinInPos()
            //{
            Pos.Byte = 0;
            Pos.Length = 0x00;
            Pos.Pos = 0;
            Pos.Mask = 0x00;
            Pos.ID = 0x00;
            //};
            byte Data = 0x00;

            switch (Msg)
            {
                case IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta:
                    Pos = CheckInPos(LHD_Drv_SHVU_SeatHtOperSta.StartByte, LHD_Drv_SHVU_SeatHtOperSta.Lenfth);
                    Pos.ID = LHD_Drv_SHVU_SeatHtOperSta.Addr;
                    break;
                case IN_CAN_LIST.RHD_Drv_SHVU_SeatHtOperSta:
                    Pos = CheckInPos(LHD_Drv_SHVU_SeatHtOperSta.StartByte, RHD_Drv_SHVU_SeatHtOperSta.Lenfth);
                    Pos.ID = RHD_Drv_SHVU_SeatHtOperSta.Addr;
                    break;
                case IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta:
                    Pos = CheckInPos(LHD_Drv_SHVU_SeatHtOperSta.StartByte, LHD_Ast_SHVU_SeatHtOperSta.Lenfth);
                    Pos.ID = LHD_Ast_SHVU_SeatHtOperSta.Addr;
                    break;
                case IN_CAN_LIST.RHD_Ast_SHVU_SeatHtOperSta:
                    Pos = CheckInPos(LHD_Drv_SHVU_SeatHtOperSta.StartByte, LHD_Drv_SHVU_SeatHtOperSta.Lenfth);
                    Pos.ID = LHD_Drv_SHVU_SeatHtOperSta.Addr;
                    break;
                case IN_CAN_LIST.C_AVNIMSButtonCmd:
                    Pos = CheckInPos(C_AVNIMSButtonCmd.StartByte, C_AVNIMSButtonCmd.Lenfth);
                    Pos.ID = C_AVNIMSButtonCmd.Addr;
                    break;
                case IN_CAN_LIST.HU_FRSeatHeatVRcmd:
                    Pos = CheckInPos(HU_FRSeatHeatVRcmd.StartByte, HU_FRSeatHeatVRcmd.Lenfth);
                    Pos.ID = HU_FRSeatHeatVRcmd.Addr;
                    break;
                case IN_CAN_LIST.HU_FLSeatHeatVRcmd:
                    Pos = CheckInPos(HU_FLSeatHeatVRcmd.StartByte, HU_FLSeatHeatVRcmd.Lenfth);
                    Pos.ID = HU_FLSeatHeatVRcmd.Addr;
                    break;
                case IN_CAN_LIST.C_FRSeatCoolerSW:
                    Pos = CheckInPos(C_FRSeatCoolerSW.StartByte, C_FRSeatCoolerSW.Lenfth);
                    Pos.ID = C_FRSeatCoolerSW.Addr;
                    break;
                case IN_CAN_LIST.C_FRSeatHeaterSW:
                    Pos = CheckInPos(C_FRSeatHeaterSW.StartByte, C_FRSeatHeaterSW.Lenfth);
                    Pos.ID = C_FRSeatHeaterSW.Addr;
                    break;
                case IN_CAN_LIST.C_FLSeatCoolerSW:
                    Pos = CheckInPos(C_FLSeatCoolerSW.StartByte, C_FLSeatCoolerSW.Lenfth);
                    Pos.ID = C_FLSeatCoolerSW.Addr;
                    break;
                case IN_CAN_LIST.C_FLSeatHeaterSW:
                    Pos = CheckInPos(C_FLSeatHeaterSW.StartByte, C_FLSeatHeaterSW.Lenfth);
                    Pos.ID = C_FLSeatHeaterSW.Addr;
                    break;
                case IN_CAN_LIST.IMS_DrvrImsSwSetSta:
                    Pos = CheckInPos(IMS_DrvrImsSwSetSta.StartByte, IMS_DrvrImsSwSetSta.Lenfth);
                    Pos.ID = IMS_DrvrImsSwSetSta.Addr;
                    break;
                case IN_CAN_LIST.IMS_DrvrImsSw1Sta:
                    Pos = CheckInPos(IMS_DrvrImsSw1Sta.StartByte, IMS_DrvrImsSw1Sta.Lenfth);
                    Pos.ID = IMS_DrvrImsSw1Sta.Addr;
                    break;
                case IN_CAN_LIST.IMS_DrvrImsSw2Sta:
                    Pos = CheckInPos(IMS_DrvrImsSw2Sta.StartByte, IMS_DrvrImsSw2Sta.Lenfth);
                    Pos.ID = IMS_DrvrImsSw2Sta.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_MemoryP1Req:
                    Pos = CheckInPos(DrvIMS_MemoryP1Req.StartByte, DrvIMS_MemoryP1Req.Lenfth);
                    Pos.ID = DrvIMS_MemoryP1Req.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_MemoryP2Req:
                    Pos = CheckInPos(DrvIMS_MemoryP2Req.StartByte, DrvIMS_MemoryP2Req.Lenfth);
                    Pos.ID = DrvIMS_MemoryP2Req.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_MemorySetEnaSta:
                    Pos = CheckInPos(DrvIMS_MemorySetEnaSta.StartByte, DrvIMS_MemorySetEnaSta.Lenfth);
                    Pos.ID = DrvIMS_MemorySetEnaSta.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_PlyBckP1Req:
                    Pos = CheckInPos(DrvIMS_PlyBckP1Req.StartByte, DrvIMS_PlyBckP1Req.Lenfth);
                    Pos.ID = DrvIMS_PlyBckP1Req.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_PlyBckP2Req:
                    Pos = CheckInPos(DrvIMS_PlyBckP2Req.StartByte, DrvIMS_PlyBckP2Req.Lenfth);
                    Pos.ID = DrvIMS_PlyBckP2Req.Addr;
                    break;
                case IN_CAN_LIST.DrvIMS_PlyBckStpReq:
                    Pos = CheckInPos(DrvIMS_PlyBckStpReq.StartByte, DrvIMS_PlyBckStpReq.Lenfth);
                    Pos.ID = DrvIMS_PlyBckStpReq.Addr;
                    break;
                case IN_CAN_LIST.SmrtIMS_PlyBckReq:
                    Pos = CheckInPos(SmrtIMS_PlyBckReq.StartByte, SmrtIMS_PlyBckReq.Lenfth);
                    Pos.ID = SmrtIMS_PlyBckReq.Addr;
                    break;
                case IN_CAN_LIST.SmrtIMS_PlyBckStpReq:
                    Pos = CheckInPos(SmrtIMS_PlyBckStpReq.StartByte, SmrtIMS_PlyBckStpReq.Lenfth);
                    Pos.ID = SmrtIMS_PlyBckStpReq.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStEsyAcsUSMSta:
                    Pos = CheckInPos(PSM_DrvStEsyAcsUSMSta.StartByte, PSM_DrvStEsyAcsUSMSta.Lenfth);
                    Pos.ID = PSM_DrvStEsyAcsUSMSta.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStSldFwdDis:
                    Pos = CheckInPos(PSM_DrvStSldFwdDis.StartByte, PSM_DrvStSldFwdDis.Lenfth);
                    Pos.ID = PSM_DrvStSldFwdDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStSldBckwdDis:
                    Pos = CheckInPos(PSM_DrvStSldBckwdDis.StartByte, PSM_DrvStSldBckwdDis.Lenfth);
                    Pos.ID = PSM_DrvStSldBckwdDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStReclneFwdDis:
                    Pos = CheckInPos(PSM_DrvStReclneFwdDis.StartByte, PSM_DrvStReclneFwdDis.Lenfth);
                    Pos.ID = PSM_DrvStReclneFwdDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStReclneBckwdDis:
                    Pos = CheckInPos(PSM_DrvStReclneBckwdDis.StartByte, PSM_DrvStReclneBckwdDis.Lenfth);
                    Pos.ID = PSM_DrvStReclneBckwdDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStTiltUpDis:
                    Pos = CheckInPos(PSM_DrvStTiltUpDis.StartByte, PSM_DrvStTiltUpDis.Lenfth);
                    Pos.ID = PSM_DrvStTiltUpDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStTiltDnDis:
                    Pos = CheckInPos(PSM_DrvStTiltDnDis.StartByte, PSM_DrvStTiltDnDis.Lenfth);
                    Pos.ID = PSM_DrvStTiltDnDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStHghtUpDis:
                    Pos = CheckInPos(PSM_DrvStHghtUpDis.StartByte, PSM_DrvStHghtUpDis.Lenfth);
                    Pos.ID = PSM_DrvStHghtUpDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStHghtDnDis:
                    Pos = CheckInPos(PSM_DrvStHghtDnDis.StartByte, PSM_DrvStHghtDnDis.Lenfth);
                    Pos.ID = PSM_DrvStHghtDnDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStLmbrDefDis:
                    Pos = CheckInPos(PSM_DrvStLmbrDefDis.StartByte, PSM_DrvStLmbrDefDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrDefDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStLmbrLoDis:
                    Pos = CheckInPos(PSM_DrvStLmbrLoDis.StartByte, PSM_DrvStLmbrLoDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrLoDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStLmbrMdlDis:
                    Pos = CheckInPos(PSM_DrvStLmbrMdlDis.StartByte, PSM_DrvStLmbrMdlDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrMdlDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStLmbrUpDis:
                    Pos = CheckInPos(PSM_DrvStLmbrUpDis.StartByte, PSM_DrvStLmbrUpDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrUpDis.Addr;
                    break;
                case IN_CAN_LIST.PSM_LmbrSigSrcTyp:
                    Pos = CheckInPos(PSM_LmbrSigSrcTyp.StartByte, PSM_LmbrSigSrcTyp.Lenfth);
                    Pos.ID = PSM_LmbrSigSrcTyp.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStReclnePosVal:
                    Pos = CheckInPos(PSM_DrvStReclnePosVal.StartByte, PSM_DrvStReclnePosVal.Lenfth);
                    Pos.ID = PSM_DrvStReclnePosVal.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStSldPosVal:
                    Pos = CheckInPos(PSM_DrvStSldPosVal.StartByte, PSM_DrvStSldPosVal.Lenfth);
                    Pos.ID = PSM_DrvStSldPosVal.Addr;
                    break;
                case IN_CAN_LIST.PSM_DrvStMvSta:
                    Pos = CheckInPos(PSM_DrvStMvSta.StartByte, PSM_DrvStMvSta.Lenfth);
                    Pos.ID = PSM_DrvStMvSta.Addr;
                    break;
                case IN_CAN_LIST.PSM_EsyAcsGetOnSta:
                    Pos = CheckInPos(PSM_EsyAcsGetOnSta.StartByte, PSM_EsyAcsGetOnSta.Lenfth);
                    Pos.ID = PSM_EsyAcsGetOnSta.Addr;
                    break;
                case IN_CAN_LIST.PSM_EsyAcsGetOffSta:
                    Pos = CheckInPos(PSM_EsyAcsGetOffSta.StartByte, PSM_EsyAcsGetOffSta.Lenfth);
                    Pos.ID = PSM_EsyAcsGetOffSta.Addr;
                    break;
                case IN_CAN_LIST.PSM_EsyAcsStpSta:
                    Pos = CheckInPos(PSM_EsyAcsStpSta.StartByte, PSM_EsyAcsStpSta.Lenfth);
                    Pos.ID = PSM_EsyAcsStpSta.Addr;
                    break;
                case IN_CAN_LIST.BCM_AccInSta:
                    Pos = CheckInPos(BCM_AccInSta.StartByte, BCM_AccInSta.Lenfth);
                    Pos.ID = BCM_AccInSta.Addr;
                    break;
                case IN_CAN_LIST.BCM_Ign1InSta:
                    Pos = CheckInPos(BCM_Ign1InSta.StartByte, BCM_Ign1InSta.Lenfth);
                    Pos.ID = BCM_Ign1InSta.Addr;
                    break;
                case IN_CAN_LIST.BCM_Ign2InSta:
                    Pos = CheckInPos(BCM_Ign2InSta.StartByte, BCM_Ign2InSta.Lenfth);
                    Pos.ID = BCM_Ign2InSta.Addr;
                    break;
                case IN_CAN_LIST.BCM_IgnSwSta:
                    Pos = CheckInPos(BCM_IgnSwSta.StartByte, BCM_IgnSwSta.Lenfth);
                    Pos.ID = BCM_IgnSwSta.Addr;
                    break;
            }

            int CanPos = -1;
            for (int i = 0; i < mCan.Can.In.Can.Send.Length; i++)
            {
                if (mCan.Can.In.Can.Send[i].ID == Pos.ID)
                {
                    CanPos = i;
                    break;
                }
            }

            if (CanPos != -1) Data = (byte)((mCan.Can.In.Can.Send[CanPos].Data[Pos.Byte] >> Pos.Pos) & Pos.Mask);

            return Data;
        }

        public enum VirtualLimiteValue
        {
            NoSetOperationAndNotLimitSet,
            NoSetOperationAndLimitSet,
            DuringSetOperationAndLotLimitSet,
            DuringSetOperationAndlimitSet
        }

        public struct VirtualLimitState
        {
            public VirtualLimiteValue Slide;
            public VirtualLimiteValue Recline;
            public VirtualLimiteValue Tilt;
            public VirtualLimiteValue Height;

            public bool OK;
        }


        public VirtualLimitState VirtualLimitDataRead()
        {
            //__CanInPos Pos;// = new __LinInPos()
            VirtualLimitState Data = new VirtualLimitState();

            //Pos.ID = 0x00;
            //Pos.ID = VirtialLimite_GetMessage.ResponseAddr;
            //Data.OK = false;

            //int CanPos = -1;
            //for (int i = 0; i < mCan.Can1.In.Can.Send.Length; i++)
            //{
            //    if (mCan.Can1.In.Can.Send[i].ID == Pos.ID)
            //    {
            //        CanPos = i;
            //        break;
            //    }
            //}

            //if (CanPos != -1)
            //{
            //    if ((mCan.Can1.In.Can.Send[CanPos].Data[0] == 0x21) && (mCan.Can1.In.Can.Send[CanPos].Data[1] == 0x71))
            //    {
            //        if ((mCan.Can1.In.Can.Send[CanPos].Data[3] == 0x12) && (mCan.Can1.In.Can.Send[CanPos].Data[4] == 0xa3))
            //        {
            //            if (mCan.Can1.In.Can.Send[CanPos].Data[2] == 0x01)
            //            {
            //                //Set
            //                Data.OK = true;
            //            }
            //            else if (mCan.Can1.In.Can.Send[CanPos].Data[2] == 0x02)
            //            {
            //                //Clear Saved
            //                Data.OK = true;
            //            }
            //        }
            //    }
            //}


            Data.OK = true;
            //Get
            if (MultiFreameData[0] == 0x01) Data.Slide = VirtualLimiteValue.NoSetOperationAndNotLimitSet;
            else if (MultiFreameData[0] == 0x02) Data.Slide = VirtualLimiteValue.NoSetOperationAndLimitSet;
            else if (MultiFreameData[0] == 0x11) Data.Slide = VirtualLimiteValue.DuringSetOperationAndLotLimitSet;
            else if (MultiFreameData[0] == 0x12) Data.Slide = VirtualLimiteValue.DuringSetOperationAndlimitSet;

            if (MultiFreameData[1] == 0x01) Data.Recline = VirtualLimiteValue.NoSetOperationAndNotLimitSet;
            else if (MultiFreameData[1] == 0x02) Data.Recline = VirtualLimiteValue.NoSetOperationAndLimitSet;
            else if (MultiFreameData[1] == 11) Data.Recline = VirtualLimiteValue.DuringSetOperationAndLotLimitSet;
            else if (MultiFreameData[1] == 12) Data.Recline = VirtualLimiteValue.DuringSetOperationAndlimitSet;

            if (MultiFreameData[2] == 0x01) Data.Tilt = VirtualLimiteValue.NoSetOperationAndLimitSet;
            else if (MultiFreameData[2] == 0x02) Data.Tilt = VirtualLimiteValue.NoSetOperationAndLimitSet;
            else if (MultiFreameData[2] == 0x11) Data.Tilt = VirtualLimiteValue.DuringSetOperationAndLotLimitSet;
            else if (MultiFreameData[2] == 0x12) Data.Tilt = VirtualLimiteValue.DuringSetOperationAndlimitSet;

            if (MultiFreameData[3] == 0x01) Data.Height = VirtualLimiteValue.NoSetOperationAndNotLimitSet;
            else if (MultiFreameData[3] == 0x02) Data.Height = VirtualLimiteValue.NoSetOperationAndLimitSet;
            else if (MultiFreameData[3] == 0x11) Data.Height = VirtualLimiteValue.DuringSetOperationAndLotLimitSet;
            else if (MultiFreameData[3] == 0x12) Data.Height = VirtualLimiteValue.DuringSetOperationAndlimitSet;
            return Data;
        }


        public __CanOutPos CheckCanOutMessage(OUT_CAN_LIST Msg, byte Data)
        {
            __CanOutPos Pos;// = new __LinPos()

            switch (Msg)
            {
                case OUT_CAN_LIST.C_AVNIMSButtonCmd:
                    Pos = CheckOutPos(C_AVNIMSButtonCmd.StartByte, C_AVNIMSButtonCmd.Lenfth);
                    Pos.ID = C_AVNIMSButtonCmd.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.HU_FRSeatHeatVRcmd:
                    Pos = CheckOutPos(HU_FRSeatHeatVRcmd.StartByte, HU_FRSeatHeatVRcmd.Lenfth);
                    Pos.ID = HU_FRSeatHeatVRcmd.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.HU_FLSeatHeatVRcmd:
                    Pos = CheckOutPos(HU_FLSeatHeatVRcmd.StartByte, HU_FLSeatHeatVRcmd.Lenfth);
                    Pos.ID = HU_FLSeatHeatVRcmd.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.C_FRSeatCoolerSW:
                    Pos = CheckOutPos(C_FRSeatCoolerSW.StartByte, C_FRSeatCoolerSW.Lenfth);
                    Pos.ID = C_FRSeatCoolerSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.C_FRSeatHeaterSW:
                    Pos = CheckOutPos(C_FRSeatHeaterSW.StartByte, C_FRSeatHeaterSW.Lenfth);
                    Pos.ID = C_FRSeatHeaterSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.C_FLSeatCoolerSW:
                    Pos = CheckOutPos(C_FLSeatCoolerSW.StartByte, C_FLSeatCoolerSW.Lenfth);
                    Pos.ID = C_FLSeatCoolerSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.C_FLSeatHeaterSW:
                    Pos = CheckOutPos(C_FLSeatHeaterSW.StartByte, C_FLSeatHeaterSW.Lenfth);
                    Pos.ID = C_FLSeatHeaterSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.IMS_DrvrImsSwSetSta:
                    Pos = CheckOutPos(IMS_DrvrImsSwSetSta.StartByte, IMS_DrvrImsSwSetSta.Lenfth);
                    Pos.ID = IMS_DrvrImsSwSetSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.IMS_DrvrImsSw1Sta:
                    Pos = CheckOutPos(IMS_DrvrImsSw1Sta.StartByte, IMS_DrvrImsSw1Sta.Lenfth);
                    Pos.ID = IMS_DrvrImsSw1Sta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.IMS_DrvrImsSw2Sta:
                    Pos = CheckOutPos(IMS_DrvrImsSw2Sta.StartByte, IMS_DrvrImsSw2Sta.Lenfth);
                    Pos.ID = IMS_DrvrImsSw2Sta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_MemoryP1Req:
                    Pos = CheckOutPos(DrvIMS_MemoryP1Req.StartByte, DrvIMS_MemoryP1Req.Lenfth);
                    Pos.ID = DrvIMS_MemoryP1Req.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_MemoryP2Req:
                    Pos = CheckOutPos(DrvIMS_MemoryP2Req.StartByte, DrvIMS_MemoryP2Req.Lenfth);
                    Pos.ID = DrvIMS_MemoryP2Req.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_MemorySetEnaSta:
                    Pos = CheckOutPos(DrvIMS_MemorySetEnaSta.StartByte, DrvIMS_MemorySetEnaSta.Lenfth);
                    Pos.ID = DrvIMS_MemorySetEnaSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_PlyBckP1Req:
                    Pos = CheckOutPos(DrvIMS_PlyBckP1Req.StartByte, DrvIMS_PlyBckP1Req.Lenfth);
                    Pos.ID = DrvIMS_PlyBckP1Req.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_PlyBckP2Req:
                    Pos = CheckOutPos(DrvIMS_PlyBckP2Req.StartByte, DrvIMS_PlyBckP2Req.Lenfth);
                    Pos.ID = DrvIMS_PlyBckP2Req.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.DrvIMS_PlyBckStpReq:
                    Pos = CheckOutPos(DrvIMS_PlyBckStpReq.StartByte, DrvIMS_PlyBckStpReq.Lenfth);
                    Pos.ID = DrvIMS_PlyBckStpReq.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.SmrtIMS_PlyBckReq:
                    Pos = CheckOutPos(SmrtIMS_PlyBckReq.StartByte, SmrtIMS_PlyBckReq.Lenfth);
                    Pos.ID = SmrtIMS_PlyBckReq.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.SmrtIMS_PlyBckStpReq:
                    Pos = CheckOutPos(SmrtIMS_PlyBckStpReq.StartByte, SmrtIMS_PlyBckStpReq.Lenfth);
                    Pos.ID = SmrtIMS_PlyBckStpReq.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStEsyAcsUSMSta:
                    Pos = CheckOutPos(PSM_DrvStEsyAcsUSMSta.StartByte, PSM_DrvStEsyAcsUSMSta.Lenfth);
                    Pos.ID = PSM_DrvStEsyAcsUSMSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStSldFwdDis:
                    Pos = CheckOutPos(PSM_DrvStSldFwdDis.StartByte, PSM_DrvStSldFwdDis.Lenfth);
                    Pos.ID = PSM_DrvStSldFwdDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStSldBckwdDis:
                    Pos = CheckOutPos(PSM_DrvStSldBckwdDis.StartByte, PSM_DrvStSldBckwdDis.Lenfth);
                    Pos.ID = PSM_DrvStSldBckwdDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStReclneFwdDis:
                    Pos = CheckOutPos(PSM_DrvStReclneFwdDis.StartByte, PSM_DrvStReclneFwdDis.Lenfth);
                    Pos.ID = PSM_DrvStReclneFwdDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStReclneBckwdDis:
                    Pos = CheckOutPos(PSM_DrvStReclneBckwdDis.StartByte, PSM_DrvStReclneBckwdDis.Lenfth);
                    Pos.ID = PSM_DrvStReclneBckwdDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStTiltUpDis:
                    Pos = CheckOutPos(PSM_DrvStTiltUpDis.StartByte, PSM_DrvStTiltUpDis.Lenfth);
                    Pos.ID = PSM_DrvStTiltUpDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStTiltDnDis:
                    Pos = CheckOutPos(PSM_DrvStTiltDnDis.StartByte, PSM_DrvStTiltDnDis.Lenfth);
                    Pos.ID = PSM_DrvStTiltDnDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStHghtUpDis:
                    Pos = CheckOutPos(PSM_DrvStHghtUpDis.StartByte, PSM_DrvStHghtUpDis.Lenfth);
                    Pos.ID = PSM_DrvStHghtUpDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStHghtDnDis:
                    Pos = CheckOutPos(PSM_DrvStHghtDnDis.StartByte, PSM_DrvStHghtDnDis.Lenfth);
                    Pos.ID = PSM_DrvStHghtDnDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStLmbrDefDis:
                    Pos = CheckOutPos(PSM_DrvStLmbrDefDis.StartByte, PSM_DrvStLmbrDefDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrDefDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStLmbrLoDis:
                    Pos = CheckOutPos(PSM_DrvStLmbrLoDis.StartByte, PSM_DrvStLmbrLoDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrLoDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStLmbrMdlDis:
                    Pos = CheckOutPos(PSM_DrvStLmbrMdlDis.StartByte, PSM_DrvStLmbrMdlDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrMdlDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStLmbrUpDis:
                    Pos = CheckOutPos(PSM_DrvStLmbrUpDis.StartByte, PSM_DrvStLmbrUpDis.Lenfth);
                    Pos.ID = PSM_DrvStLmbrUpDis.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_LmbrSigSrcTyp:
                    Pos = CheckOutPos(PSM_LmbrSigSrcTyp.StartByte, PSM_LmbrSigSrcTyp.Lenfth);
                    Pos.ID = PSM_LmbrSigSrcTyp.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStReclnePosVal:
                    Pos = CheckOutPos(PSM_DrvStReclnePosVal.StartByte, PSM_DrvStReclnePosVal.Lenfth);
                    Pos.ID = PSM_DrvStReclnePosVal.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStSldPosVal:
                    Pos = CheckOutPos(PSM_DrvStSldPosVal.StartByte, PSM_DrvStSldPosVal.Lenfth);
                    Pos.ID = PSM_DrvStSldPosVal.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_DrvStMvSta:
                    Pos = CheckOutPos(PSM_DrvStMvSta.StartByte, PSM_DrvStMvSta.Lenfth);
                    Pos.ID = PSM_DrvStMvSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_EsyAcsGetOnSta:
                    Pos = CheckOutPos(PSM_EsyAcsGetOnSta.StartByte, PSM_EsyAcsGetOnSta.Lenfth);
                    Pos.ID = PSM_EsyAcsGetOnSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_EsyAcsGetOffSta:
                    Pos = CheckOutPos(PSM_EsyAcsGetOffSta.StartByte, PSM_EsyAcsGetOffSta.Lenfth);
                    Pos.ID = PSM_EsyAcsGetOffSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.PSM_EsyAcsStpSta:
                    Pos = CheckOutPos(PSM_EsyAcsStpSta.StartByte, PSM_EsyAcsStpSta.Lenfth);
                    Pos.ID = PSM_EsyAcsStpSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.BCM_AccInSta:
                    Pos = CheckOutPos(BCM_AccInSta.StartByte, BCM_AccInSta.Lenfth);
                    Pos.ID = BCM_AccInSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.BCM_Ign1InSta:
                    Pos = CheckOutPos(BCM_Ign1InSta.StartByte, BCM_Ign1InSta.Lenfth);
                    Pos.ID = BCM_Ign1InSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.BCM_Ign2InSta:
                    Pos = CheckOutPos(BCM_Ign2InSta.StartByte, BCM_Ign2InSta.Lenfth);
                    Pos.ID = BCM_Ign2InSta.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_CAN_LIST.BCM_IgnSwSta:
                    Pos = CheckOutPos(BCM_IgnSwSta.StartByte, BCM_IgnSwSta.Lenfth);
                    Pos.ID = BCM_IgnSwSta.Addr;
                    Pos.Data = Data;
                    break;
                default:
                    Pos.Byte = 0;
                    Pos.Data = 0x00;
                    Pos.Pos = 0;
                    Pos.Mask = 0x00;
                    Pos.ID = 0x00;
                    break;
            }

            return Pos;
        }

        public __CanOutMessage CheckCanOutMessage(OUT_CAN_LIST Msg, bool OnOff)
        {
            __CanOutMessage Pos = new __CanOutMessage()
            {
                Data = new byte[8],
                ID = 0
            };

            switch (Msg)
            {
                case OUT_CAN_LIST.PSeat_Slide_Fwd:
                    Pos.ID = PSeat_Slide_Fwd.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Slide_Fwd.MoveData;
                    else Pos.Data = PSeat_Slide_Fwd.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Slide_Bwd:
                    Pos.ID = PSeat_Slide_Bwd.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Slide_Bwd.MoveData;
                    else Pos.Data = PSeat_Slide_Bwd.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Recline_Fwd:
                    Pos.ID = PSeat_Recline_Fwd.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Recline_Fwd.MoveData;
                    else Pos.Data = PSeat_Recline_Fwd.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Recline_Bwd:
                    Pos.ID = PSeat_Recline_Bwd.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Recline_Bwd.MoveData;
                    else Pos.Data = PSeat_Recline_Bwd.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Tilt_Up:
                    Pos.ID = PSeat_Tilt_Up.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Tilt_Up.MoveData;
                    else Pos.Data = PSeat_Tilt_Up.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Tilt_Down:
                    Pos.ID = PSeat_Tilt_Down.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Tilt_Down.MoveData;
                    else Pos.Data = PSeat_Tilt_Down.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Height_Up:
                    Pos.ID = PSeat_Height_Up.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Height_Up.MoveData;
                    else Pos.Data = PSeat_Height_Up.StopData;
                    break;
                case OUT_CAN_LIST.PSeat_Height_Down:
                    Pos.ID = PSeat_Height_Down.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = PSeat_Height_Down.MoveData;
                    else Pos.Data = PSeat_Height_Down.StopData;
                    break;
                case OUT_CAN_LIST.VirtialLimite_All:
                    Pos.ID = VirtialLimite_All.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = VirtialLimite_All.LimitSet;
                    else Pos.Data = VirtialLimite_All.Clear;
                    break;
                case OUT_CAN_LIST.VirtialLimite_Slide:
                    Pos.ID = VirtialLimite_Slide.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = VirtialLimite_Slide.LimitSet;
                    else Pos.Data = VirtialLimite_Slide.Clear;
                    break;
                case OUT_CAN_LIST.VirtialLimite_Recline:
                    Pos.ID = VirtialLimite_Recline.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = VirtialLimite_Recline.LimitSet;
                    else Pos.Data = VirtialLimite_Recline.Clear;
                    break;
                case OUT_CAN_LIST.VirtialLimite_Tilt:
                    Pos.ID = VirtialLimite_Tilt.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = VirtialLimite_Tilt.LimitSet;
                    else Pos.Data = VirtialLimite_Tilt.Clear;
                    break;
                case OUT_CAN_LIST.VirtialLimite_Height:
                    Pos.ID = VirtialLimite_Height.RequestAddr;
                    if (OnOff == true)
                        Pos.Data = VirtialLimite_Height.LimitSet;
                    else Pos.Data = VirtialLimite_Height.Clear;
                    break;
                case OUT_CAN_LIST.VirtialLimite_GetMessage:
                    Pos.ID = VirtialLimite_GetMessage.RequestAddr;
                    Pos.Data = VirtialLimite_GetMessage.GetRequest;

                    break;
                default:
                    Pos.ID = -1;
                    break;
            }

            return Pos;
        }

        private __CanInPos CheckInPos(int Start, int Length)
        {
            __CanInPos Pos = new __CanInPos()
            {
                Byte = 0,
                Length = 0x00,
                Pos = 0,
                Mask = 0x00,
                ID = 0x00
            };

            Pos.Byte = GetStartByte((short)Start);
            Pos.Pos = GetStartPos((short)Start);
            Pos.Mask = GetMasking((short)Length);
            Pos.Length = (byte)Length;
            return Pos;
        }

        private __CanOutPos CheckOutPos(int Start, int Length)
        {
            __CanOutPos Pos = new __CanOutPos()
            {
                Byte = 0,
                Data = 0x00,
                Pos = 0,
                Mask = 0x00,
                ID = 0x00
            };


            Pos.Byte = GetStartByte((short)Start);
            Pos.Pos = GetStartPos((short)Start);
            Pos.Mask = GetMasking((short)Length);

            return Pos;
        }

        private short GetStartByte(short value)
        {
            //get
            //{
            //    return (short)lStartByteToIn;
            //}
            //set
            //{
            //    lStartByteToIn = (short)(value / 8);
            //}
            return (short)(value / 8);
        }

        private short GetStartPos(short value)
        {
            //get
            //{
            //    return (short)lStartPosToIn;
            //}
            //set
            //{
            //    lStartByteToIn = (short)(value % 8);
            //}
            return (short)(value % 8);
        }

        private byte GetMasking(short value)
        {
            //get
            //{
            //    return lMaskingToIn;
            //}

            //set
            //{
            byte lMaskingToIn = 0x00;
            switch (value)
            {
                case 1: lMaskingToIn = 0x01; break;
                case 2: lMaskingToIn = 0x03; break;
                case 3: lMaskingToIn = 0x07; break;
                case 4: lMaskingToIn = 0x0f; break;
                case 5: lMaskingToIn = 0x1f; break;
                case 6: lMaskingToIn = 0x3f; break;
                case 7: lMaskingToIn = 0x7f; break;
                case 8: lMaskingToIn = 0xff; break;
                default: lMaskingToIn = 0x00; break;
            }
            //}
            return lMaskingToIn;
        }

        //private bool LH_HeaterHIghLamp;
        //private bool LH_HeaterMidLamp;
        //private bool LH_HeaterLowLamp;
        //private bool RH_HeaterHIghLamp;
        //private bool RH_HeaterMidLamp;
        //private bool RH_HeaterLowLamp;
        //private bool LH_VentHIghLamp;
        //private bool LH_VentMidLamp;
        //private bool LH_VentLowLamp;
        //private bool RH_VentHIghLamp;
        //private bool RH_VentMidLamp;
        //private bool RH_VentLowLamp;

        private void CheckLinIn()
        {
            LH_HeaterHighLamp = CheckLinInMessage(IN_LIN_LIST.LH_HeaterSWLEDHigh) == (byte)LH_HeaterSWLEDHigh.Data.On ? true : false;
            LH_HeaterMidLamp = CheckLinInMessage(IN_LIN_LIST.LH_HeaterSWLEDMid) == (byte)LH_HeaterSWLEDMid.Data.On ? true : false;
            LH_HeaterLowLamp = CheckLinInMessage(IN_LIN_LIST.LH_HeaterSWLEDLow) == (byte)LH_HeaterSWLEDLow.Data.On ? true : false;
            RH_HeaterHighLamp = CheckLinInMessage(IN_LIN_LIST.RH_HeaterSWLEDHigh) == (byte)RH_HeaterSWLEDHigh.Data.On ? true : false;
            RH_HeaterMidLamp = CheckLinInMessage(IN_LIN_LIST.RH_HeaterSWLEDMid) == (byte)RH_HeaterSWLEDMid.Data.On ? true : false;
            RH_HeaterLowLamp = CheckLinInMessage(IN_LIN_LIST.RH_HeaterSWLEDLow) == (byte)RH_HeaterSWLEDLow.Data.On ? true : false;

            LH_VentHighLamp = CheckLinInMessage(IN_LIN_LIST.LH_VentSWLEDHigh) == (byte)LH_VentSWLEDHigh.Data.On ? true : false;
            LH_VentMidLamp = CheckLinInMessage(IN_LIN_LIST.LH_VentSWLEDMid) == (byte)LH_VentSWLEDMid.Data.On ? true : false;
            LH_VentLowLamp = CheckLinInMessage(IN_LIN_LIST.LH_VentSWLEDLow) == (byte)LH_VentSWLEDLow.Data.On ? true : false;
            RH_VentHighLamp = CheckLinInMessage(IN_LIN_LIST.RH_VentSWLEDHigh) == (byte)RH_VentSWLEDHigh.Data.On ? true : false;
            RH_VentMidLamp = CheckLinInMessage(IN_LIN_LIST.RH_VentSWLEDMid) == (byte)RH_VentSWLEDMid.Data.On ? true : false;
            RH_VentLowLamp = CheckLinInMessage(IN_LIN_LIST.RH_VentSWLEDLow) == (byte)RH_VentSWLEDLow.Data.On ? true : false;
            return;
        }

        private void CheckCanIn()
        {
            bool LHD;
            bool RHD;
            bool Flag;

            LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.Heater_VentOff ? true : false;
            RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.Heater_VentOff ? true : false;

            if ((LHD == true) && (RHD == true))
            {
                LH_HeaterHighLamp = false;
                LH_HeaterMidLamp = false;
                LH_HeaterLowLamp = false;

                LH_VentHighLamp = false;
                LH_VentMidLamp = false;
                LH_VentLowLamp = false;
            }
            else
            {
                Flag = false;

                //------------ 3단
                LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.HeatherHigh ? true : false;
                RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.HeatherHigh ? true : false;

                if ((LHD == true) || (RHD == true))
                {
                    LH_HeaterHighLamp = true;
                    LH_HeaterMidLamp = true;
                    LH_HeaterLowLamp = true;
                    Flag = true;
                }

                if (Flag == false)
                {
                    //------------ 2단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.HeatherMid ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.HeatherMid ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        LH_HeaterHighLamp = false;
                        LH_HeaterMidLamp = true;
                        LH_HeaterLowLamp = true;
                        Flag = true;
                    }
                }

                if (Flag == false)
                {
                    //------------ 1단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.HeaterLow ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.HeaterLow ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        LH_HeaterHighLamp = false;
                        LH_HeaterMidLamp = false;
                        LH_HeaterLowLamp = true;
                        Flag = true;
                    }
                }

                //통풍
                Flag = false;

                //------------ 3단
                LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.VentHigh ? true : false;
                RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.VentHigh ? true : false;

                if ((LHD == true) || (RHD == true))
                {
                    LH_VentHighLamp = true;
                    LH_VentMidLamp = true;
                    LH_VentLowLamp = true;
                    Flag = true;
                }

                if (Flag == false)
                {
                    //------------ 2단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.VentMid ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.VentMid ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        LH_VentHighLamp = false;
                        LH_VentMidLamp = true;
                        LH_VentLowLamp = true;
                        Flag = true;
                    }
                }

                if (Flag == false)
                {
                    //------------ 1단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)LHD_Drv_SHVU_SeatHtOperSta.Data.VentLow ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Drv_SHVU_SeatHtOperSta) == (byte)RHD_Drv_SHVU_SeatHtOperSta.Data.VentLow ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        LH_VentHighLamp = false;
                        LH_VentMidLamp = false;
                        LH_VentLowLamp = true;
                        Flag = true;
                    }
                }
            }


            LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.Heater_VentOff ? true : false;
            RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.Heater_VentOff ? true : false;

            if ((LHD == true) && (RHD == true))
            {
                RH_HeaterHighLamp = false;
                RH_HeaterMidLamp = false;
                RH_HeaterLowLamp = false;

                RH_VentHighLamp = false;
                RH_VentMidLamp = false;
                RH_VentLowLamp = false;
            }
            else
            {
                Flag = false;

                //------------ 3단
                LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.HeatherHigh ? true : false;
                RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.HeatherHigh ? true : false;

                if ((LHD == true) || (RHD == true))
                {
                    RH_HeaterHighLamp = true;
                    RH_HeaterMidLamp = true;
                    RH_HeaterLowLamp = true;
                    Flag = true;
                }

                if (Flag == false)
                {
                    //------------ 2단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.HeatherMid ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.HeatherMid ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        RH_HeaterHighLamp = false;
                        RH_HeaterMidLamp = true;
                        RH_HeaterLowLamp = true;
                        Flag = true;
                    }
                }

                if (Flag == false)
                {
                    //------------ 1단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.HeaterLow ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.HeaterLow ? true : false;

                    if ((RHD == true) || (RHD == true))
                    {
                        RH_HeaterHighLamp = false;
                        RH_HeaterMidLamp = false;
                        RH_HeaterLowLamp = true;
                        Flag = true;
                    }
                }

                //통풍
                Flag = false;

                //------------ 3단
                LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.VentHigh ? true : false;
                RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.VentHigh ? true : false;

                if ((LHD == true) || (RHD == true))
                {
                    RH_VentHighLamp = true;
                    RH_VentMidLamp = true;
                    RH_VentLowLamp = true;
                    Flag = true;
                }

                if (Flag == false)
                {
                    //------------ 2단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.VentMid ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.VentMid ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        RH_VentHighLamp = false;
                        RH_VentMidLamp = true;
                        RH_VentLowLamp = true;
                        Flag = true;
                    }
                }

                if (Flag == false)
                {
                    //------------ 1단
                    LHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)LHD_Ast_SHVU_SeatHtOperSta.Data.VentLow ? true : false;
                    RHD = CheckCanInMessage(IN_CAN_LIST.LHD_Ast_SHVU_SeatHtOperSta) == (byte)RHD_Ast_SHVU_SeatHtOperSta.Data.VentLow ? true : false;

                    if ((LHD == true) || (RHD == true))
                    {
                        RH_VentHighLamp = false;
                        RH_VentMidLamp = false;
                        RH_VentLowLamp = true;
                        Flag = true;
                    }
                }
            }

            return;
        }

        public bool LH_HeaterHighLamp { get; set; }
        public bool LH_HeaterMidLamp { get; set; }
        public bool LH_HeaterLowLamp { get; set; }
        public bool RH_HeaterHighLamp { get; set; }
        public bool RH_HeaterMidLamp { get; set; }
        public bool RH_HeaterLowLamp { get; set; }
        public bool LH_VentHighLamp { get; set; }
        public bool LH_VentMidLamp { get; set; }
        public bool LH_VentLowLamp { get; set; }
        public bool RH_VentHighLamp { get; set; }
        public bool RH_VentMidLamp { get; set; }
        public bool RH_VentLowLamp { get; set; }

        public byte CheckLinInMessage(IN_LIN_LIST Msg)
        {
            __CanInPos Pos;// = new __LinInPos()
            //{
            Pos.Byte = 0;
            Pos.Length = 0x00;
            Pos.Pos = 0;
            Pos.Mask = 0x00;
            Pos.ID = 0x00;
            //};
            byte Data = 0x00;

            switch (Msg)
            {
                case IN_LIN_LIST.SWEvenLEDDimReq:
                    Pos = CheckInPos(SWEvenLEDDimReq.StartByte, SWEvenLEDDimReq.Lenfth);
                    Pos.ID = SWEvenLEDDimReq.Addr;
                    break;
                case IN_LIN_LIST.RH_VentSWLEDMid:
                    Pos = CheckInPos(RH_VentSWLEDMid.StartByte, RH_VentSWLEDMid.Lenfth);
                    Pos.ID = RH_VentSWLEDMid.Addr;
                    break;
                case IN_LIN_LIST.LH_VentSWLEDHigh:
                    Pos = CheckInPos(LH_VentSWLEDHigh.StartByte, LH_VentSWLEDHigh.Lenfth);
                    Pos.ID = LH_VentSWLEDHigh.Addr;
                    break;
                case IN_LIN_LIST.LH_VentSWLEDLow:
                    Pos = CheckInPos(LH_VentSWLEDLow.StartByte, LH_VentSWLEDLow.Lenfth);
                    Pos.ID = LH_VentSWLEDLow.Addr;
                    break;
                case IN_LIN_LIST.RH_VentSWLEDHigh:
                    Pos = CheckInPos(RH_VentSWLEDHigh.StartByte, RH_VentSWLEDHigh.Lenfth);
                    Pos.ID = RH_VentSWLEDHigh.Addr;
                    break;
                case IN_LIN_LIST.RH_VentSWLEDLow:
                    Pos = CheckInPos(RH_VentSWLEDLow.StartByte, RH_VentSWLEDLow.Lenfth);
                    Pos.ID = RH_VentSWLEDLow.Addr;
                    break;
                case IN_LIN_LIST.LH_VentSWLEDMid:
                    Pos = CheckInPos(LH_VentSWLEDMid.StartByte, LH_VentSWLEDMid.Lenfth);
                    Pos.ID = LH_VentSWLEDMid.Addr;
                    break;
                case IN_LIN_LIST.SWOddLEDDimReq:
                    Pos = CheckInPos(SWOddLEDDimReq.StartByte, SWOddLEDDimReq.Lenfth);
                    Pos.ID = SWOddLEDDimReq.Addr;
                    break;
                case IN_LIN_LIST.LH_HeaterSWLEDMid:
                    Pos = CheckInPos(LH_HeaterSWLEDMid.StartByte, LH_HeaterSWLEDMid.Lenfth);
                    Pos.ID = LH_HeaterSWLEDMid.Addr;
                    break;
                case IN_LIN_LIST.RH_HeaterSWLEDHigh:
                    Pos = CheckInPos(RH_HeaterSWLEDHigh.StartByte, RH_HeaterSWLEDHigh.Lenfth);
                    Pos.ID = RH_HeaterSWLEDHigh.Addr;
                    break;
                case IN_LIN_LIST.RH_HeaterSWLEDLow:
                    Pos = CheckInPos(RH_HeaterSWLEDLow.StartByte, RH_HeaterSWLEDLow.Lenfth);
                    Pos.ID = RH_HeaterSWLEDLow.Addr;
                    break;
                case IN_LIN_LIST.LH_HeaterSWLEDLow:
                    Pos = CheckInPos(LH_HeaterSWLEDLow.StartByte, LH_HeaterSWLEDLow.Lenfth);
                    Pos.ID = LH_HeaterSWLEDLow.Addr;
                    break;
                case IN_LIN_LIST.LH_HeaterSWLEDHigh:
                    Pos = CheckInPos(LH_HeaterSWLEDHigh.StartByte, LH_HeaterSWLEDHigh.Lenfth);
                    Pos.ID = LH_HeaterSWLEDHigh.Addr;
                    break;
                case IN_LIN_LIST.RH_HeaterSWLEDMid:
                    Pos = CheckInPos(RH_HeaterSWLEDMid.StartByte, RH_HeaterSWLEDMid.Lenfth);
                    Pos.ID = RH_HeaterSWLEDMid.Addr;
                    break;
            }

            int CanPos = -1;
            for (int i = 0; i < mCan.Lin.In.Lin.Send.Length; i++)
            {
                if (mCan.Lin.In.Lin.Send[i].ID == Pos.ID)
                {
                    CanPos = i;
                    break;
                }
            }

            if (CanPos != -1) Data = (byte)((mCan.Lin.In.Lin.Send[CanPos].Data[Pos.Byte] >> Pos.Pos) & Pos.Mask);

            return Data;
        }

        public __CanOutPos CheckLinOutMessage(OUT_LIN_LIST Msg, byte Data)
        {
            __CanOutPos Pos = new __CanOutPos()
            {
                Data = 0x00,
                ID = 0,
                Byte = 0,
                Mask = 0,
                Pos = 0
            };

            switch (Msg)
            {
                case OUT_LIN_LIST.LH_HeaterSW:
                    Pos = CheckOutPos(LH_HeaterSW.StartByte, LH_HeaterSW.Lenfth);
                    Pos.ID = LH_HeaterSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.RH_HeaterSW:
                    Pos = CheckOutPos(RH_HeaterSW.StartByte, RH_HeaterSW.Lenfth);
                    Pos.ID = RH_HeaterSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.LH_VentSW:
                    Pos = CheckOutPos(LH_VentSW.StartByte, LH_VentSW.Lenfth);
                    Pos.ID = LH_VentSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.RH_VentSW:
                    Pos = CheckOutPos(RH_VentSW.StartByte, RH_VentSW.Lenfth);
                    Pos.ID = RH_VentSW.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.LH_HeaterSWRaw:
                    Pos = CheckOutPos(LH_HeaterSWRaw.StartByte, LH_HeaterSWRaw.Lenfth);
                    Pos.ID = LH_HeaterSWRaw.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.RH_HeaterSWRaw:
                    Pos = CheckOutPos(RH_HeaterSWRaw.StartByte, RH_HeaterSWRaw.Lenfth);
                    Pos.ID = RH_HeaterSWRaw.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.SWEvenLEDDimRes:
                    Pos = CheckOutPos(SWEvenLEDDimRes.StartByte, SWEvenLEDDimRes.Lenfth);
                    Pos.ID = SWEvenLEDDimRes.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.SWLINError:
                    Pos = CheckOutPos(SWLINError.StartByte, SWLINError.Lenfth);
                    Pos.ID = SWLINError.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.LH_VentSWRaw:
                    Pos = CheckOutPos(LH_VentSWRaw.StartByte, LH_VentSWRaw.Lenfth);
                    Pos.ID = LH_VentSWRaw.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.RH_VentSWRaw:
                    Pos = CheckOutPos(RH_VentSWRaw.StartByte, RH_VentSWRaw.Lenfth);
                    Pos.ID = RH_VentSWRaw.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.SWOddLEDDimRes:
                    Pos = CheckOutPos(SWOddLEDDimRes.StartByte, SWOddLEDDimRes.Lenfth);
                    Pos.ID = SWOddLEDDimRes.Addr;
                    Pos.Data = Data;
                    break;
                case OUT_LIN_LIST.SWLEDStatus:
                    Pos = CheckOutPos(SWLEDStatus.StartByte, SWLEDStatus.Lenfth);
                    Pos.ID = SWLEDStatus.Addr;
                    Pos.Data = Data;
                    break;
                default:
                    Pos.ID = -1;
                    break;
            }

            return Pos;
        }

        private BackgroundWorker backgroundWorker1 = null;
        private void ThreadSetting()
        {
            backgroundWorker1 = new BackgroundWorker();

            //ReportProgress메소드를 호출하기 위해서 반드시 true로 설정, false일 경우 ReportProgress메소드를 호출하면 exception 발생
            backgroundWorker1.WorkerReportsProgress = true;
            //스레드에서 취소 지원 여부
            backgroundWorker1.WorkerSupportsCancellation = true;
            //스레드가 run시에 호출되는 핸들러 등록
            backgroundWorker1.DoWork += new DoWorkEventHandler(BackgroundWorker1_DoWork);
            // ReportProgress메소드 호출시 호출되는 핸들러 등록
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker1_ProgressChanged);
            // 스레드 완료(종료)시 호출되는 핸들러 동록
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker1_RunWorkerCompleted);


            // 스레드가 Busy(즉, run)가 아니라면
            if (backgroundWorker1.IsBusy != true)
            {
                // 스레드 작동!! 아래 함수 호출 시 위에서 bw.DoWork += new DoWorkEventHandler(bw_DoWork); 에 등록한 핸들러가
                // 호출 됩니다.

                backgroundWorker1.RunWorkerAsync();
            }
            return;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //바로 위에서 worker.ReportProgress((i * 10));호출 시 
            // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 된다고
            // 하였는데요.. 이 부분에서는 기존 Thread에서 처럼 Dispatcher를 이용하지 않아도 됩니다. 
            // 즉 아래처럼!!사용이 가능합니다.
            //this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");

            // 기존의 Thread클래스에서 아래와 같이 UI 엘리먼트를 갱신하려면
            // Dispatcher.BeginInvoke(delegate() 
            // {
            //        this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");
            // )};
            //처럼 처리해야 할 것입니다. 그러나 바로 UI 엘리먼트를 업데이트 하고 있죠??
        }


        //스레드의 run함수가 종료될 경우 해당 핸들러가 호출됩니다.
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            //스레드가 종료한 이유(사용자 취소, 완료, 에러)에 맞쳐 처리하면 됩니다.
            if ((e.Cancelled == true))
            {
            }
            else if (!(e.Error == null))
            {

            }
            else
            {

            }
        }

        //private short runCh = 0;
        //private bool EStopFlagClearFlag = false;
        //private bool EStopClearFlag = false;
        //private int LinOutPos = 0;
        //private int LinInPos = 0;
        //private int CanOutPos1 = 0;
        private int CanOutPos = 0;
        private bool PowerOnToFirst1 = false;
        private bool PowerOnToFirst2 = false;
        //private int CanInPos = 0;

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;


            //LinOutPos = 0;
            //LinInPos = 0;
            CanOutPos = 0;
            //CanOutPos2 = 0;
            //CanInPos = 0;

            do
            {
                //CancellationPending 속성이 true로 set되었다면(위에서 CancelAsync 메소드 호출 시 true로 set된다고 하였죠?
                if ((worker.CancellationPending == true))
                {
                    //루프를 break한다.(즉 스레드 run 핸들러를 벗어나겠죠)
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // 이곳에는 스레드에서 처리할 연산을 넣으시면 됩니다.

                    Processing();

                    Thread.Sleep(1);
                    // 스레드 진행상태 보고 - 이 메소드를 호출 시 위에서 
                    // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 됩니다.
                    worker.ReportProgress(10);
                }
                if (mControl.isExit == true)
                {
                    worker.CancelAsync();
                }
            } while (true);
            //while (ExitFlag == false);
        }

        private void Processing()
        {
            //LinInOut();
            CanInOut();
            CheckCanIn();
            return;
        }

        //private void LinInOut()
        //{
        //    if (mControl.GetLin != null)
        //    {
        //        if (mControl.GetLin.isOpen(mControl.GetLinChannel) == true)
        //        {
        //            if (LinOutPos < 0) LinOutPos = 0;
        //            if (mCan.Lin.Out.Lin.Max <= LinOutPos) LinOutPos = 0;

        //            if (0 < mCan.Lin.Out.Lin.Send[LinOutPos].ID)
        //            {
        //                mCan.Lin.Out.Lin.Send[LinOutPos].last = mControl.공용함수.timeGetTimems();

        //                if (mCan.Lin.Out.Lin.Send[LinOutPos].sendtime == 0) mCan.Lin.Out.Lin.Send[LinOutPos].sendtime = 100;
        //                if (mCan.Lin.Out.Lin.Send[LinOutPos].sendtime <= (mCan.Lin.Out.Lin.Send[LinOutPos].last - mCan.Lin.Out.Lin.Send[LinOutPos].first))
        //                {
        //                    __CanMsg Msg = new __CanMsg() { DATA = new byte[8] };
        //                    Array.Clear(Msg.DATA, 0, 8);
        //                    Msg.ID = 0;
        //                    Msg.Length = 0;

        //                    if (0 < mCan.Lin.Out.Lin.Send[LinOutPos].Length)
        //                    {
        //                        Msg.ID = mCan.Lin.Out.Lin.Send[LinOutPos].ID;
        //                        Msg.Length = mCan.Lin.Out.Lin.Send[LinOutPos].Length;
        //                        Array.Copy(mCan.Lin.Out.Lin.Send[LinOutPos].Data, 0, Msg.DATA, 0, 8);
        //                    }

        //                    mControl.GetLin.LinWrite(mControl.GetLinChannel, Msg, LinControl.Direction.dirPublisher, LinControl.ChecksumType.cstEnhanced);
        //                    mCan.Lin.Out.Lin.Send[LinOutPos].first = mControl.공용함수.timeGetTimems();
        //                }
        //            }
        //            LinOutPos++;
        //            //------------------------------------------------------------------------------
        //            //if (LinInPos < 0) LinInPos = 0;
        //            //if (mCan.Lin.In.Lin.Max <= LinInPos) LinInPos = 0;

        //            //if (0 < mCan.Lin.In.Lin.Send[LinInPos].ID)
        //            //{
        //            //    mCan.Lin.In.Lin.Send[LinInPos].last = mControl.공용함수.timeGetTimems();

        //            //    if (mCan.Lin.In.Lin.Send[LinInPos].sendtime == 0) mCan.Lin.In.Lin.Send[LinInPos].sendtime = 100;
        //            //    if (mCan.Lin.In.Lin.Send[LinInPos].sendtime <= (mCan.Lin.In.Lin.Send[LinInPos].last - mCan.Lin.In.Lin.Send[LinInPos].first))
        //            //    {
        //            //        __CanMsg Msg = new __CanMsg() { DATA = new byte[8] };
        //            //        Array.Clear(Msg.DATA, 0, 8);
        //            //        Msg.ID = 0;
        //            //        Msg.Length = 0;

        //            //        if (0 < mCan.Lin.In.Lin.Send[LinInPos].Length)
        //            //        {
        //            //            Msg.ID = mCan.Lin.In.Lin.Send[LinInPos].ID;
        //            //            Msg.Length = mCan.Lin.In.Lin.Send[LinInPos].Length;
        //            //            if (Msg.ID == 0x1f)
        //            //                Msg.Length = 4;
        //            //            else if (Msg.ID == 0x3c)
        //            //                Msg.Length = 8;
        //            //            else Msg.Length = 8;

        //            //            Array.Copy(mCan.Lin.In.Lin.Send[LinInPos].Data, 0, Msg.DATA, 0, 8);
        //            //        }

        //            //        mControl.GetLin.LinWrite(mControl.GetLinChannel, Msg, LinControl.Direction.dirSubscriberAutoLength, LinControl.ChecksumType.cstAuto);
        //            //        mCan.Lin.In.Lin.Send[LinInPos].first = mControl.공용함수.timeGetTimems();
        //            //    }
        //            //}
        //            //LinInPos++;

        //            //------------------------------------------------------------------------------
        //            mControl.GetLin.ReadMessages(mControl.GetLinChannel);
        //            //Lin.ReadMultiMessage();

        //            if (mControl.GetLin.isData(mControl.GetLinChannel) == true)
        //            {
        //                __CanMsg Msg = mControl.GetLin.GetLin(mControl.GetLinChannel);

        //                if (0 <= Msg.ID)
        //                {
        //                    if (Msg.ID == 0x08)
        //                    {
        //                        byte Even = (byte)((Msg.DATA[0] & 0x80) >> 3);
        //                        byte Odd = (byte)((Msg.DATA[1] & 0x80) >> 5);

        //                        byte Data1 = (byte)(mCan.Lin.Out.Lin.Send[0].Data[1] & ~0x04);
        //                        byte Data2 = (byte)(mCan.Lin.Out.Lin.Send[0].Data[1] & ~0x40);

        //                        mCan.Lin.Out.Lin.Send[0].Data[1] = (byte)(Data1 | Even | Data2);

        //                        __CanMsg wMsg = new __CanMsg()
        //                        {
        //                            DATA = new byte[8],
        //                            Length = 2,
        //                            ID = mCan.Lin.Out.Lin.Send[0].ID
        //                        };

        //                        Array.Copy(mCan.Lin.Out.Lin.Send[0].Data, 0, wMsg.DATA, 0, 8);
        //                        mControl.GetLin.LinWrite(mControl.GetLinChannel, wMsg, LinControl.Direction.dirPublisher, LinControl.ChecksumType.cstEnhanced);
        //                    }

        //                    bool Flag = false;
        //                    for (int i = 2; i < mCan.Lin.In.Lin.Max; i++)
        //                    {
        //                        if (mCan.Lin.In.Lin.Send[i].ID == Msg.ID)
        //                        {
        //                            mCan.Lin.In.Lin.Send[i].Length = Msg.Length;
        //                            Array.Copy(Msg.DATA, 0, mCan.Lin.In.Lin.Send[i].Data, 0, Msg.Length);
        //                            Flag = true;
        //                        }
        //                    }

        //                    if (Flag == false)
        //                    {
        //                        if (mCan.Lin.In.Lin.Max < 20)
        //                        {
        //                            if (mCan.Lin.In.Lin.Send[mCan.Lin.In.Lin.Max].ID == Msg.ID)
        //                            {
        //                                mCan.Lin.In.Lin.Send[mCan.Lin.In.Lin.Max].Length = Msg.Length;
        //                                Array.Copy(Msg.DATA, 0, mCan.Lin.In.Lin.Send[mCan.Lin.In.Lin.Max].Data, 0, Msg.Length);
        //                                mCan.Lin.In.Lin.Max++;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return;
        //}

        private bool VersionReadFlag = false;
        private string SwVerData = "";
        public void SendVersionReadMessage()
        {
            SwVerData = null;
            VersionReadFlag = true;
            return;
        }
        public string GetSWVersion
        {
            get
            {
                return SwVerData;
            }
        }

        private bool PSeatCanWriteFlag = false;
        private long PSeatCanWriteFirst;
        private long PSeatCanWriteLast;
        private long PSeatCanWriteFirst2;
        private long PSeatCanWriteLast2;
        private bool PowerOnFlag = false;
        private int MultiFrameLength = 0;
        private int PowerOnFirstSendToCount = 0;


        public bool GetCanFail
        {
            get
            {
                return CanFailFlag;
            }
        }
        private void CanInOut()
        {
            if (mControl.GetPower.isPowerOn == true)
            {
                PowerOnFlag = true;
            }
            else
            {
                if (PowerOnFlag == true) PowerOnFlag = false;
                if (PowerOnToFirst1 == false) PowerOnToFirst1 = true;
                if (PowerOnToFirst2 == false) PowerOnToFirst2 = true;
                if (PowerOnFirstSendToCount != 0) PowerOnFirstSendToCount = 0;

                PSeatCanWriteFirst2 = mControl.공용함수.timeGetTimems();
                PSeatCanWriteLast2 = mControl.공용함수.timeGetTimems();
            }

            if (PowerOnFlag == false) return;

            if (mControl.GetCan != null)
            {
                //if (mControl.GetCan.isOpen(0) == true)
                //{
                //    if (CanOutPos1 < 0) CanOutPos1 = 0;
                //    if (mCan.Can1.Out.Can.Max <= CanOutPos1) CanOutPos1 = 0;

                //    if (0 < mCan.Can1.Out.Can.Send[CanOutPos1].ID)
                //    {
                //        mCan.Can1.Out.Can.Send[CanOutPos1].last = mControl.공용함수.timeGetTimems();

                //        if (mCan.Can1.Out.Can.Send[CanOutPos1].sendtime == 0) mCan.Can1.Out.Can.Send[CanOutPos1].sendtime = 100;
                //        if (mCan.Can1.Out.Can.Send[CanOutPos1].sendtime <= (mCan.Can1.Out.Can.Send[CanOutPos1].last - mCan.Can1.Out.Can.Send[CanOutPos1].first))
                //        {
                //            __CanMsg oMsg = new __CanMsg() { DATA = new byte[8] };
                //            Array.Clear(oMsg.DATA, 0, 8);
                //            oMsg.ID = 0;
                //            oMsg.Length = 0;

                //            if (0 < mCan.Can1.Out.Can.Send[CanOutPos1].Length)
                //            {
                //                oMsg.ID = mCan.Can1.Out.Can.Send[CanOutPos1].ID;
                //                oMsg.Length = mCan.Can1.Out.Can.Send[CanOutPos1].Length;
                //                Array.Copy(mCan.Can1.Out.Can.Send[CanOutPos1].Data, 0, oMsg.DATA, 0, 8);
                //            }

                //            mControl.GetCan.WriteCan(0, oMsg, false);
                //            mCan.Can1.Out.Can.Send[CanOutPos1].first = mControl.공용함수.timeGetTimems();
                //        }
                //    }
                //    CanOutPos1++;

                //    //------------------------------------------------------------------------------
                //    __CanMsg Msg = mControl.GetCan.ReadCan(0, false);
                //    //Can.ReadMultiMessage();


                //    if ((0 < Msg.Length) && (0 <= Msg.ID))
                //    {
                //        bool InFlag = true;

                //        if (VirtialLimite_GetMessage.ResponseAddr == Msg.ID)
                //        {
                //            if ((Msg.DATA[0] == 0x10) && (Msg.DATA[1] == 0x71) && (Msg.DATA[2] != 0x03)) // Get
                //                InFlag = false;
                //            else if (MultiFrameRead == true)
                //                InFlag = false;
                //            else InFlag = true;
                //        }
                //        else
                //        {
                //            InFlag = true;
                //        }


                //        if (InFlag == true)
                //        {
                //            bool Flag = false;
                //            for (int i = 0; i < mCan.Can1.In.Can.Max; i++)
                //            {
                //                if (mCan.Can1.In.Can.Send[i].ID == Msg.ID)
                //                {
                //                    mCan.Can1.In.Can.Send[i].Length = Msg.Length;
                //                    Array.Copy(Msg.DATA, 0, mCan.Can1.In.Can.Send[i].Data, 0, Msg.Length);
                //                    Flag = true;
                //                }
                //            }

                //            if (Flag == false)
                //            {
                //                if (mCan.Can1.In.Can.Max < 20)
                //                {
                //                    if (mCan.Can1.In.Can.Send[mCan.Can1.In.Can.Max].ID == Msg.ID)
                //                    {
                //                        mCan.Can1.In.Can.Send[mCan.Can1.In.Can.Max].Length = Msg.Length;
                //                        Array.Copy(Msg.DATA, 0, mCan.Can1.In.Can.Send[mCan.Can1.In.Can.Max].Data, 0, Msg.Length);
                //                        mCan.Can1.In.Can.Max++;
                //                    }
                //                }
                //            }
                //        }

                //        if (InFlag == false)
                //        {
                //            if (VirtialLimite_GetMessage.ResponseAddr == Msg.ID)
                //            {
                //                if ((Msg.DATA[0] == 0x10) && (Msg.DATA[1] == 0x71) && (Msg.DATA[2] == 0x03)) // Get
                //                {
                //                    if ((Msg.DATA[3] == 0x12) && (Msg.DATA[4] == 0xa3))
                //                    {
                //                        Array.Clear(MultiFreameData, 0, 100);
                //                        MultiFreameDataPos = 0;
                //                        MultiFrameRead = true;

                //                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[5];
                //                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[6];
                //                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[7];


                //                        __CanMsg sMsg = new __CanMsg() { DATA = new byte[8] };

                //                        mControl.공용함수.timedelay(50);
                //                        sMsg.Length = 8;
                //                        sMsg.ID = VirtialLimite_GetMessage.RequestAddr;
                //                        sMsg.DATA[0] = 0x30;
                //                        sMsg.DATA[1] = 0x08;
                //                        sMsg.DATA[2] = 0x0a;
                //                        sMsg.DATA[3] = 0x00;
                //                        sMsg.DATA[4] = 0x00;
                //                        sMsg.DATA[5] = 0x00;
                //                        sMsg.DATA[6] = 0x00;
                //                        sMsg.DATA[7] = 0x00;
                //                        mControl.GetCan.WriteCan(0, sMsg, false);
                //                    }
                //                }
                //                else if (Msg.DATA[0] == 0x21)
                //                {
                //                    if (MultiFrameRead == true)
                //                    {
                //                        for (int i = 1; i < 8; i++)
                //                        {
                //                            MultiFreameData[MultiFreameDataPos++] = Msg.DATA[i];
                //                        }
                //                        MultiFrameRead = false;
                //                    }
                //                }
                //            }
                //        }
                //    }
                //}

                if (mControl.GetCan.isOpen(0) == true)
                {
                    //if (CanOutPos2 < 0) CanOutPos2 = 0;
                    //if (mCan.Can2.Out.Can.Max <= CanOutPos2) CanOutPos2 = 0;

                    //if (0 < mCan.Can2.Out.Can.Send[CanOutPos2].ID)
                    //{
                    //    mCan.Can2.Out.Can.Send[CanOutPos2].last = mControl.공용함수.timeGetTimems();

                    //    if (mCan.Can2.Out.Can.Send[CanOutPos2].sendtime == 0) mCan.Can2.Out.Can.Send[CanOutPos2].sendtime = 100;
                    //    if (mCan.Can2.Out.Can.Send[CanOutPos2].sendtime <= (mCan.Can2.Out.Can.Send[CanOutPos2].last - mCan.Can2.Out.Can.Send[CanOutPos2].first))
                    //    {
                    //        __CanMsg oMsg = new __CanMsg() { DATA = new byte[8] };
                    //        Array.Clear(oMsg.DATA, 0, 8);
                    //        oMsg.ID = 0;
                    //        oMsg.Length = 0;

                    //        if (0 < mCan.Can2.Out.Can.Send[CanOutPos2].Length)
                    //        {
                    //            oMsg.ID = mCan.Can2.Out.Can.Send[CanOutPos2].ID;
                    //            oMsg.Length = mCan.Can2.Out.Can.Send[CanOutPos2].Length;
                    //            Array.Copy(mCan.Can2.Out.Can.Send[CanOutPos2].Data, 0, oMsg.DATA, 0, 8);
                    //        }

                    //        mControl.GetCan.WriteCan(1, oMsg, true);
                    //        mCan.Can2.Out.Can.Send[CanOutPos2].first = mControl.공용함수.timeGetTimems();
                    //    }
                    //}
                    //



                    if (PSeatCanWriteFlag == false)
                    {
                        if (PSeatDataSendFlag == false)
                        {
                            if (PowerOnFlag == true)
                            {
                                __CanMsg oMsg = new __CanMsg() { DATA = new byte[8] };
                                Array.Clear(oMsg.DATA, 0, 8);
                                oMsg.ID = 0;
                                oMsg.Length = 0;


                                if (PowerOnToFirst1 == true)
                                {
                                    PSeatCanWriteLast2 = mControl.공용함수.timeGetTimems();
                                    if (400 <= (PSeatCanWriteLast2 - PSeatCanWriteFirst2))
                                    {
                                        if (PowerOnFirstSendToCount < 5)
                                        {
                                            oMsg.ID = PSeatControlStart.ID;
                                            oMsg.Length = PSeatControlStart.Length;
                                            Array.Copy(PSeatControlStart.DATA, 0, oMsg.DATA, 0, 8);

                                            PowerOnFirstSendToCount++;
                                            if (5 <= PowerOnFirstSendToCount) PowerOnToFirst1 = false;

                                            PSeatCanWriteFirst2 = mControl.공용함수.timeGetTimems();
                                            PSeatCanWriteLast2 = mControl.공용함수.timeGetTimems();
                                        }
                                        else
                                        {
                                            PowerOnToFirst1 = false;
                                        }
                                    }
                                }
                                else if (PowerOnToFirst2 == true)
                                {
                                    PowerOnToFirst2 = false;
                                    oMsg.ID = PSeatControlStart.ID;
                                    oMsg.Length = VirtialLimite_CommunicationStart.Length;
                                    Array.Copy(VirtialLimite_CommunicationStart.DATA, 0, oMsg.DATA, 0, 8);
                                }
                                else
                                {
                                    //if (CanOutPos2 == 0)
                                    //{
                                    //    oMsg.ID = MotorStatusRead.ID;
                                    //    oMsg.Length = MotorStatusRead.Length;
                                    //    Array.Copy(MotorStatusRead.DATA, 0, oMsg.DATA, 0, 8);
                                    //}
                                    //else if (CanOutPos2 == 1)
                                    //{
                                    //    oMsg.ID = PositionRead.ID;
                                    //    oMsg.Length = PositionRead.Length;
                                    //    Array.Copy(PositionRead.DATA, 0, oMsg.DATA, 0, 8);
                                    //}
                                    //else
                                    //{
                                    //    oMsg.ID = PSeatSwRead.ID;
                                    //    oMsg.Length = PositionRead.Length;
                                    //    Array.Copy(PositionRead.DATA, 0, oMsg.DATA, 0, 8);
                                    //}

                                    if (VersionReadFlag == false)
                                    {
                                        oMsg.ID = VirtialLimite_GetMessage.RequestAddr;
                                        oMsg.Length = 8;
                                        Array.Copy(VirtialLimite_GetMessage.GetRequest, 0, oMsg.DATA, 0, 8);
                                    }
                                    else
                                    {
                                        VersionReadFlag = false;
                                        oMsg.ID = SWVersionRead.RequestAddr;
                                        oMsg.Length = 8;
                                        Array.Copy(SWVersionRead.GetRequest, 0, oMsg.DATA, 0, 8);
                                    }

                                }
                                mControl.GetCan.WriteCan(0, oMsg, false);
                            }
                        }
                        //CanOutPos2++;
                        PSeatCanWriteFirst = mControl.공용함수.timeGetTimems();
                        PSeatCanWriteLast = mControl.공용함수.timeGetTimems();
                        PSeatCanWriteFlag = true;
                    }


                    //------------------------------------------------------------------------------
                    __CanMsg Msg = mControl.GetCan.ReadCan(0, false);
                    //Can.ReadMultiMessage();

                    if (Msg.ID != -1)
                    {
                        CanFailFlag = false;
                        CanFailCheckFirstTime = mControl.공용함수.timeGetTimems();
                        CanFailCheckLastTime = mControl.공용함수.timeGetTimems();
                    }
                    else
                    {
                        CanFailCheckLastTime = mControl.공용함수.timeGetTimems();
                        if (1500 <= (CanFailCheckLastTime - CanFailCheckFirstTime))
                        {
                            CanFailFlag = true;
                        }
                    }

                    if (Msg.ID == 0x7ab) //PSeat
                    {

                        if ((Msg.DATA[1] == 0x62) && (Msg.DATA[2] == 0xf1) && (Msg.DATA[3] == 0xa0))
                        {
                            string s = Encoding.Default.GetString(Msg.DATA, 4, 3);
                            SwVerData = s.Substring(0, 1);
                            SwVerData += ".";
                            SwVerData += s.Substring(1, 2);
                        }
                        else
                        {
                            if (Msg.DATA[0] == 0x10)
                            {
                                if ((Msg.DATA[4] == 0x12) && (Msg.DATA[5] == 0xa3))
                                {
                                    //Array.Clear(MultiFreameData, 0, 20);
                                    MultiFreameDataPos = 0;
                                    MultiFrameRead = true;
                                    MultiFrameLength = (int)Msg.DATA[1];
                                    //MultiFreameData[MultiFreameDataPos++] = Msg.DATA[5];
                                    MultiFreameData[MultiFreameDataPos++] = Msg.DATA[6];
                                    MultiFreameData[MultiFreameDataPos++] = Msg.DATA[7];


                                    __CanMsg sMsg = new __CanMsg() { DATA = new byte[8] };

                                    mControl.공용함수.timedelay(50);
                                    sMsg.Length = 8;
                                    sMsg.ID = VirtialLimite_GetMessage.RequestAddr;
                                    sMsg.DATA[0] = 0x30;
                                    sMsg.DATA[1] = 0x08;
                                    sMsg.DATA[2] = 0x0a;
                                    sMsg.DATA[3] = 0x00;
                                    sMsg.DATA[4] = 0x00;
                                    sMsg.DATA[5] = 0x00;
                                    sMsg.DATA[6] = 0x00;
                                    sMsg.DATA[7] = 0x00;
                                    mControl.GetCan.WriteCan(0, sMsg, false);
                                }
                            }
                            else if (Msg.DATA[0] == 0x21)
                            {
                                if (MultiFrameRead == true)
                                {
                                    for (int i = 1; i < 8; i++)
                                    {
                                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[i];
                                        if (MultiFrameLength <= MultiFreameDataPos) break;
                                    }

                                    MultiFrameRead = false;
                                }

                                if (MultiFrameLength <= MultiFreameDataPos)
                                {
                                    CanOutPos++;
                                    PSeatCanWriteFlag = false;
                                    if (2 < CanOutPos)
                                    {
                                        CanOutPos = 0;
                                    }
                                }
                            }
                            else if (Msg.DATA[0] == 0x22)
                            {
                                if (MultiFrameRead == true)
                                {
                                    for (int i = 1; i < 8; i++)
                                    {
                                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[i];
                                        if (MultiFrameLength <= MultiFreameDataPos) break;
                                    }

                                    MultiFrameRead = false;
                                }

                                CanOutPos++;
                                PSeatCanWriteFlag = false;
                                if (2 < CanOutPos)
                                {
                                    CanOutPos = 0;
                                }
                            }
                        }
                    }
                    if (PSeatCanWriteFlag == true)
                    {
                        PSeatCanWriteLast = mControl.공용함수.timeGetTimems();
                        if (1000 <= (PSeatCanWriteLast - PSeatCanWriteFirst))
                        {
                            CanOutPos++;
                            PSeatCanWriteFlag = false;
                            if (2 < CanOutPos)
                            {
                                CanOutPos = 0;
                            }
                        }
                    }
                    if ((0 < Msg.Length) && (0 <= Msg.ID))
                    {
                        bool InFlag = true;

                        if (VirtialLimite_GetMessage.ResponseAddr == Msg.ID)
                        {
                            if ((Msg.DATA[0] == 0x10) && (Msg.DATA[1] == 0x71) && (Msg.DATA[2] != 0x03)) // Get
                                InFlag = false;
                            else if (MultiFrameRead == true)
                                InFlag = false;
                            else InFlag = true;
                        }
                        else
                        {
                            InFlag = true;
                        }


                        if (InFlag == true)
                        {
                            bool Flag = false;
                            for (int i = 0; i < mCan.Can.In.Can.Max; i++)
                            {
                                if (mCan.Can.In.Can.Send[i].ID == Msg.ID)
                                {
                                    mCan.Can.In.Can.Send[i].Length = Msg.Length;
                                    Array.Copy(Msg.DATA, 0, mCan.Can.In.Can.Send[i].Data, 0, Msg.Length);
                                    Flag = true;
                                }
                            }

                            if (Flag == false)
                            {
                                if (mCan.Can.In.Can.Max < 20)
                                {
                                    if (mCan.Can.In.Can.Send[mCan.Can.In.Can.Max].ID == Msg.ID)
                                    {
                                        mCan.Can.In.Can.Send[mCan.Can.In.Can.Max].Length = Msg.Length;
                                        Array.Copy(Msg.DATA, 0, mCan.Can.In.Can.Send[mCan.Can.In.Can.Max].Data, 0, Msg.Length);
                                        mCan.Can.In.Can.Max++;
                                    }
                                }
                            }
                        }

                        if (InFlag == false)
                        {
                            if (VirtialLimite_GetMessage.ResponseAddr == Msg.ID)
                            {
                                if ((Msg.DATA[0] == 0x10) && (Msg.DATA[1] == 0x71) && (Msg.DATA[2] == 0x03)) // Get
                                {
                                    if ((Msg.DATA[3] == 0x12) && (Msg.DATA[4] == 0xa3))
                                    {
                                        Array.Clear(MultiFreameData, 0, 100);
                                        MultiFreameDataPos = 0;
                                        MultiFrameRead = true;

                                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[5];
                                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[6];
                                        MultiFreameData[MultiFreameDataPos++] = Msg.DATA[7];


                                        __CanMsg sMsg = new __CanMsg() { DATA = new byte[8] };

                                        mControl.공용함수.timedelay(50);
                                        sMsg.Length = 8;
                                        sMsg.ID = VirtialLimite_GetMessage.RequestAddr;
                                        sMsg.DATA[0] = 0x30;
                                        sMsg.DATA[1] = 0x08;
                                        sMsg.DATA[2] = 0x0a;
                                        sMsg.DATA[3] = 0x00;
                                        sMsg.DATA[4] = 0x00;
                                        sMsg.DATA[5] = 0x00;
                                        sMsg.DATA[6] = 0x00;
                                        sMsg.DATA[7] = 0x00;
                                        mControl.GetCan.WriteCan(0, sMsg, false);
                                    }
                                }
                                else if (Msg.DATA[0] == 0x21)
                                {
                                    if (MultiFrameRead == true)
                                    {
                                        for (int i = 1; i < 8; i++)
                                        {
                                            MultiFreameData[MultiFreameDataPos++] = Msg.DATA[i];
                                        }
                                        MultiFrameRead = false;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return;
        }
        private byte[] MultiFreameData = new byte[20];
        private short MultiFreameDataPos = 0;
        private bool MultiFrameRead = false;
        //private short MultiFreameDataLength = 0;

        private bool PSeatDataSendFlag = false;

        public void CanDataOutput(OUT_CAN_LIST Msg, byte Data)
        {
            if ((Msg != OUT_CAN_LIST.PSeat_Slide_Fwd) && (Msg != OUT_CAN_LIST.PSeat_Slide_Bwd) && (Msg != OUT_CAN_LIST.PSeat_Recline_Fwd) && (Msg != OUT_CAN_LIST.PSeat_Recline_Bwd) &&
                (Msg != OUT_CAN_LIST.PSeat_Tilt_Up) && (Msg != OUT_CAN_LIST.PSeat_Tilt_Down) && (Msg != OUT_CAN_LIST.PSeat_Height_Up) && (Msg != OUT_CAN_LIST.PSeat_Height_Down) &&
                (Msg != OUT_CAN_LIST.VirtialLimite_All) && (Msg != OUT_CAN_LIST.VirtialLimite_Slide) && (Msg != OUT_CAN_LIST.VirtialLimite_Recline) && (Msg != OUT_CAN_LIST.VirtialLimite_Tilt) &&
                (Msg != OUT_CAN_LIST.VirtialLimite_Height))
            {
                __CanOutPos Pos = CheckCanOutMessage(Data: Data, Msg: Msg);

                short LPos = -1;
                for (short i = 0; i < mCan.Can.Out.Can.Max; i++)
                {
                    if (LPos == -1)
                    {
                        if (Pos.ID == mCan.Can.Out.Can.Send[i].ID)
                        {
                            LPos = i;
                            break;
                        }
                    }
                }

                if (LPos == -1) return;
                //Masking
                byte Mastking = (byte)~((Pos.Mask << Pos.Pos) & 0xff);
                byte LData = (byte)(mCan.Can.Out.Can.Send[LPos].Data[Pos.Byte] & Mastking);

                //Data
                LData |= (byte)((Pos.Data << Pos.Pos) & 0xff);
                mCan.Can.Out.Can.Send[LPos].Data[Pos.Byte] = LData;


                __CanMsg LMsg = new __CanMsg()
                {
                    DATA = new byte[8],
                    ID = 0,
                    Length = 0
                };

                LMsg.ID = mCan.Can.Out.Can.Send[LPos].ID;
                LMsg.Length = mCan.Can.Out.Can.Send[LPos].Length;
                Array.Copy(mCan.Can.Out.Can.Send[LPos].Data, 0, LMsg.DATA, 0, 8);

                mControl.GetCan.WriteCan(0, LMsg, false);
                mCan.Can.Out.Can.Send[LPos].first = mControl.공용함수.timeGetTimems();
            }
            else
            {
                PSeatDataSendFlag = true;
                mControl.공용함수.timedelay(100);
                if ((Msg != OUT_CAN_LIST.VirtialLimite_All) && (Msg != OUT_CAN_LIST.VirtialLimite_Slide) && (Msg != OUT_CAN_LIST.VirtialLimite_Recline) && (Msg != OUT_CAN_LIST.VirtialLimite_Tilt) &&
                    (Msg != OUT_CAN_LIST.VirtialLimite_Height))
                {
                    __CanOutMessage Pos = CheckCanOutMessage(OnOff: Data == 0x00 ? false : true, Msg: Msg);

                    short LPos = -1;
                    for (short i = 0; i < mCan.Can.Out.Can.Max; i++)
                    {
                        if (LPos == -1)
                        {
                            if (Pos.ID == mCan.Can.Out.Can.Send[i].ID)
                            {
                                LPos = i;
                                break;
                            }
                        }
                    }

                    if (LPos == -1) return;

                    mCan.Can.Out.Can.Send[LPos].Data = Pos.Data;

                    __CanMsg LMsg = new __CanMsg()
                    {
                        DATA = new byte[8],
                        ID = 0,
                        Length = 0
                    };

                    LMsg.ID = mCan.Can.Out.Can.Send[LPos].ID;
                    LMsg.Length = mCan.Can.Out.Can.Send[LPos].Length;
                    //Array.Copy(mCan.Can1.Out.Can.Send[LPos].Data, 0, LMsg.DATA, 0, 8);
                    LMsg.DATA = mCan.Can.Out.Can.Send[LPos].Data;

                    mControl.GetCan.WriteCan(0, LMsg, false);
                    mCan.Can.Out.Can.Send[LPos].first = mControl.공용함수.timeGetTimems();
                }
                else
                {
                    __CanOutMessage Pos = CheckCanOutMessage(OnOff: Data == 0x00 ? false : true, Msg: Msg);

                    short LPos = -1;
                    for (short i = 0; i < mCan.Can.Out.Can.Max; i++)
                    {
                        if (LPos == -1)
                        {
                            if (Pos.ID == mCan.Can.Out.Can.Send[i].ID)
                            {
                                LPos = i;
                                break;
                            }
                        }
                    }

                    if (LPos == -1) return;

                    mCan.Can.Out.Can.Send[LPos].Data = Pos.Data;

                    __CanMsg LMsg = new __CanMsg()
                    {
                        DATA = new byte[8],
                        ID = 0,
                        Length = 0
                    };

                    LMsg.ID = mCan.Can.Out.Can.Send[LPos].ID;
                    LMsg.Length = mCan.Can.Out.Can.Send[LPos].Length;
                    //Array.Copy(mCan.Can1.Out.Can.Send[LPos].Data, 0, LMsg.DATA, 0, 8);
                    LMsg.DATA = mCan.Can.Out.Can.Send[LPos].Data;

                    mControl.GetCan.WriteCan(0, LMsg, false);
                    mCan.Can.Out.Can.Send[LPos].first = mControl.공용함수.timeGetTimems();
                }
                mControl.공용함수.timedelay(100);
                PSeatDataSendFlag = false;
            }
            return;
        }

        //public void LINDataOutput(short Ch, OUT_LIN_LIST Msg, byte Data)
        //{

        //    __CanOutPos Pos = CheckLinOutMessage(Data: Data, Msg: Msg);

        //    short LPos = -1;
        //    for (short i = 0; i < mCan.Lin.Out.Lin.Max; i++)
        //    {
        //        if (LPos == -1)
        //        {
        //            if (Pos.ID == mCan.Lin.Out.Lin.Send[i].ID)
        //            {
        //                LPos = i;
        //                break;
        //            }
        //        }
        //    }

        //    if (LPos == -1) return;
        //    //Masking
        //    byte Mastking = (byte)~((Pos.Mask << Pos.Pos) & 0xff);
        //    byte LData = (byte)(mCan.Lin.Out.Lin.Send[LPos].Data[Pos.Byte] & Mastking);

        //    //Data
        //    LData |= (byte)((Pos.Data << Pos.Pos) & 0xff);
        //    mCan.Lin.Out.Lin.Send[LPos].Data[Pos.Byte] = LData;


        //    __CanMsg LMsg = new __CanMsg()
        //    {
        //        DATA = new byte[8],
        //        ID = 0,
        //        Length = 0
        //    };

        //    LMsg.ID = mCan.Lin.Out.Lin.Send[LPos].ID;
        //    LMsg.Length = mCan.Lin.Out.Lin.Send[LPos].Length;
        //    Array.Copy(mCan.Lin.Out.Lin.Send[LPos].Data, 0, LMsg.DATA, 0, 8);

        //    mControl.GetLin.LinWrite(Ch, LMsg);
        //    mCan.Lin.Out.Lin.Send[LPos].first = mControl.공용함수.timeGetTimems();
        //    return;
        //}
    }
}
