using System;
using System.Drawing;
using System.Runtime.InteropServices;
using MES;


namespace PSeatTester
{
    [StructLayout(LayoutKind.Explicit)]
    public struct union_r
    {
        [FieldOffset(0)]
        public int i;
        [FieldOffset(0)]
        public byte c1;
        [FieldOffset(1)]
        public byte c2;
        [FieldOffset(2)]
        public byte c3;
        [FieldOffset(3)]
        public byte c4;
    };
      
  
    public enum MENU
    {
        NONE,
        AGING_TESTING,
        PERFORMANCE_TESTING,
        AGING_SETTING,
        PERFORMANCE_SETTING,
        OPTION,
        LIN1,
        LIN2,
        CAN,
        PASSWORD,
        AGING_DATAVIEW,
        PERFORMANCE_DATAVIEW,
        SELF
    }


    public class RESULT
    {
        public const short READY = 0;
        public const short PASS = 1;
        public const short NG = 2;
        public const short REJECT = 2;
        public const short END = 3;
        public const short STOP = 4;
        public const short CLEAR = 5;
        public const short TEST = 6;
        public const short NOT_TEST = 7;
        public const short TEST_STANDBY = 8;        
    };


    public class IO_IN
    {
        public const short PASS = 0;
        public const short RESET = 1;
        public const short RH_SELECT = 2;
        public const short SEAT_POWER = 3;
        public const short LUMBER_2WAY = 5; //new 
        public const short AUTO = 4;
                
        public const short LHD_RHD = 6; //new 

        public const short SLIDE_FWD = 8;
        public const short SLIDE_BWD = 9;
        public const short RECLINE_FWD = 10;
        public const short RECLINE_BWD = 11;
        public const short TILT_UP = 12;
        public const short TILT_DN = 13;
        public const short HEIGHT_UP = 14;
        public const short HEIGHT_DN = 15;

        public const short LUMBER_FWD = 17;
        public const short LUMBER_BWD = 16;
        public const short LUMBER_UP = 18;
        public const short LUMBER_DN = 19;
        public const short SlideDeliveryPosSensor = 20;
        public const short ReclineDeliveryPosSensor = 21;
        public const short PRODUCT = 22;
        public const short JIG_UP = 23;
        public const short PIN_CONNECTION_FWD = 24;
        public const short PIN_CONNECTION_BWD = 25;

        public const short IMS_SET_SW = 30;
        public const short IMS_M1_SW = 29;
        public const short IMS_M2_SW = 28;
        public const short OPTION_WALKIN = 27; //new 
        public const short OPTION_IMS= 26;
        public const short PIN_CONNECTION_SW = 7;
    }

    public class IO_IN2
    {
        public const short PASS = 0;
        public const short RESET = 1;
        public const short LHD = 3;
        public const short RHD = 2;
        public const short LH_SELECT = 4;
        public const short RH_SELECT = 5;

        public const short OPTION_IMS = 6;
        public const short OPTION_WALKIN = 7;        
        public const short OPTION_POWER = 8;

        public const short LUMBER_4WAY = 9; //new 
        public const short LUMBER_2WAY = 10; //new 
        public const short LUMBER_0WAY = 11; //new 

        public const short AUTO = 13;
        public const short MANUAL = 12;

        public const short SLIDE_FWD = 18;
        public const short SLIDE_BWD = 17;
        public const short RECLINE_FWD = 20;
        public const short RECLINE_BWD = 19;
        public const short TILT_UP = 22;
        public const short TILT_DN = 21;
        public const short HEIGHT_UP = 24;
        public const short HEIGHT_DN = 23;

        public const short LUMBER_FWD = 26;
        public const short LUMBER_BWD = 25;
        public const short LUMBER_UP = 28;
        public const short LUMBER_DN = 27;
        //public const short SlideDeliveryPosSensor = 20;
        //public const short ReclineDeliveryPosSensor = 21;
        //public const short PRODUCT = 22;
        //public const short JIG_UP = 23;
        //public const short PIN_CONNECTION_FWD = 24;
        //public const short PIN_CONNECTION_BWD = 25;

        public const short IMS_SET_SW = 14;
        public const short IMS_M1_SW = 15;
        public const short IMS_M2_SW = 16;
        public const short PIN_CONNECTION_SW = 28;
    }

    /*
    IMS - LHD - LH (운전석) - 12WAY
              - RH (조수석) - 8WAY
              
          RHD - RH (운전석) - 12WAY
              - LH (조수석) - 8WAY


    POWER - LHD - LH (운전석) - 12WAY , 10 WAY
                - RH (조수석) - 8WAY, WALK IN
              
            RHD - RH (운전석) - 12WAY , 10 WAY
                - LH (조수석) - 8WAY, WALK IN
    */

    public class IO_OUT
    {
        public const short RED_LAMP = 3; //0
        public const short YELLOW_LAMP = 2; //1
        public const short GREEN_LAMP = 1; //2
        public const short BUZZER = 0; // 3

        public const short PRODUCT_IN = 4;
        public const short TEST_OK = 5;
        //public const short TEST_JIG_DOWN = 6;
        public const short DOOR_OPEN = 6;
        public const short TEST_ING = 7;

        //public const short JIG_ORG = 8;
        public const short PIN_CONNECTION = 24;
    }

    public enum LH_RH
    {
        LH,
        RH
    }
    public enum PSEAT_SELECT
    {
        PSEAT_8WAY_4W,
        PSEAT_8WAY,
        PSEAT_8WAY_2W,
        PSEAT_8WAY_WALKIN
    }
    
    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __MinMax__
    {
        public double Min;
        public double Max;
    }

    public struct __MinMaxToInt__
    {
        public int Min;
        public int Max;
    }

    public struct __Port__
    {
        public string Port;
        public int Speed;
    }


    public struct __LinDevice__
    {
        public short Device;
        public int Speed;
    }
    public struct __CanDevice__
    {
        public short Device;
        public short Channel;
        public short ID;
        public int Speed;
    }

    public struct __TcpIP__
    {
        public string IP;
        public int Port;
    }

    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __Config__
    {
        /// <summary>
        /// Master
        /// </summary>
        //public __LinDevice__ Lin;
        public __CanDevice__ Can;
        public __Port__ NoiseMeter;
        public __Port__ SmartIOPort;
        public __TCPIP__ Client;
        public __TCPIP__ Server;
        public __TcpIP__ Board;
        public __TcpIP__ PC;
        public __Port__ Power;
        public __Port__ Panel;
        public short BattID;
        public short CurrID;
        public float PinConnectionDelay;
        public bool AutoConnection;
        public bool IMSSeatToCanControl;
        public float KalmanSpeed;
        public bool UseSmartIO;
    }
    
    public struct __Time__
    {
        public int Hour;
        public short Min;
        public short Sec;
        public short mSec;
    }


    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    //[StructLayout(LayoutKind.Sequential, Pack = 1)]    
    public struct __CanData__
    {
        public short Data;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 50)]
        public char[] Title; //[50];
    };

    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __ItemCan__
    {
        public short StartBit;
        public short Size;
        public short Mode; // 0이면 일반데이타가 1 이면 숫치 데이타가 들어가는 항목임 2 이면 아스키 데이타가 들어간다.
        public short DataCounter;
        public short Length;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 50)]
        public char[] Title; //[50];
        public int CanID;
        public short S_ID;
        public short ReceiveTime; // 전송 간격을 갖는다.
        public bool CanLin; // true can, false Lin
        public bool InOut; // true 이면 output mode

        [MarshalAs(UnmanagedType.ByValArray, ArraySubType = UnmanagedType.Struct, SizeConst = 40)]
        public __CanData__[] Data;//[40];
    }

    public struct __Can__
    {
        public short ItemCounter;

        [MarshalAs(UnmanagedType.ByValArray, ArraySubType = UnmanagedType.Struct, SizeConst = 700)]
        public __ItemCan__[] Item; // [700]
    };
    public struct __sCan__
    {
        public bool Run;
        public short sID;
        public short dBit;
    }

    public struct __CanMsg
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
        public byte[] DATA;// [8];
        public int Length;
        public int ID;
    }


    public struct __SendCan__
    {
        public int ID;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
        public byte[] Data; //[8]
        public int Length;
        public long first;
        public long last;
        public long sendtime;
        //public byte AliveCnt;
    }

    public struct __InOutCanMsg__
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 20)]
        public __SendCan__[] Send; // [20]
        public int Max;
    }

    public struct CanInOutStruct
    {
        public __InOutCanMsg__ Can;
    }
    public struct LinInOutStruct
    {
        public __InOutCanMsg__ Lin;
    }

    public struct __InOutCan__
    {
        public CanInOutStruct In;
        public CanInOutStruct Out;
    }
    public struct __InOutLin__
    {
        public LinInOutStruct In;
        public LinInOutStruct Out;
    }


    public struct __IOData__
    {
        public short Card;
        public short Pos;
        public byte Data;
    }

    public struct __LinOutPos
    {
        /// <summary>
        /// 데이타
        /// </summary>
        public byte Data;
        /// <summary>
        /// 초기화 데이타 
        /// </summary>
        public byte Mask;
        //시작 위치
        public short Byte;
        /// <summary>
        /// 비트 위치
        /// </summary>
        public short Pos;
        /// <summary>
        /// Lin FID/PID
        /// </summary>
        public byte ID;
    }
    public struct __LinInPos
    {
        /// <summary>
        /// 데이타
        /// </summary>
        public byte Length;
        /// <summary>
        /// 초기화 데이타 
        /// </summary>
        public byte Mask;
        //시작 위치
        public short Byte;
        /// <summary>
        /// 비트 위치
        /// </summary>
        public short Pos;
        /// <summary>
        /// Lin FID/PID
        /// </summary>
        public byte ID;
    }

    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __TestDataItem__
    {
        public short Result;
        public float Data;
        public bool Test;
    }

    public struct __TestSoundData__
    {
        public short StartResult;
        public short RunResult;
        public float StartData;
        public float RunData;
        public bool Test;
    }


    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __TestCanLinDataItem__
    {
        public short Result;
        public byte Data;
        public short Message;
        public bool Test;
    }

    [Serializable, StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]//CharSet = CharSet.Unicode를 선언해 주지 않으면 한글 처리할 때 파일에 저장하거나 할 경우 에러가 발생한다.
    public struct __TestData__
    {
        public short Result;
        //[MarshalAs(UnmanagedType.ByValArray, SizeConst = 50)]
        //public char[] Time;
        public string Time;

        public __TestDataItem__ Ims;
        public __TestDataItem__ Slide;
        public __TestDataItem__ Recline;
        public __TestDataItem__ Tilt;
        public __TestDataItem__ Height;
        public __TestSoundData__ SoundSlide;
        public __TestSoundData__ SoundRecline;
        public __TestSoundData__ SoundTilt;
        public __TestSoundData__ SoundHeight;
        public __TestSoundData__ SoundLumberUpDn;
        public __TestSoundData__ SoundLumberFwdBwd;


        public __TestDataItem__ LumberFwdBwd;
        public __TestDataItem__ LumberUpDn;

        public __TestDataItem__ SlideFwdTime;
        public __TestDataItem__ SlideBwdTime;
        public __TestDataItem__ ReclineFwdTime;
        public __TestDataItem__ ReclineBwdTime;
        public __TestDataItem__ TiltUpTime;
        public __TestDataItem__ TiltDnTime;
        public __TestDataItem__ HeightUpTime;
        public __TestDataItem__ HeightDnTime;
    }

    public struct __Infor__ 
    {
        public string Date;
        public string DataName;
        public int TotalCount;
        public int OkCount;
        public int NgCount;
        public bool ReBootingFlag;
    }
    public struct __CanLin__
    {
        public __InOutCan__ Can;
        public __InOutLin__ Lin;
    }
    public struct __CanInPos
    {
        /// <summary>
        /// 데이타
        /// </summary>
        public byte Length;
        /// <summary>
        /// 초기화 데이타 
        /// </summary>
        public byte Mask;
        //시작 위치
        public short Byte;
        /// <summary>
        /// 비트 위치
        /// </summary>
        public short Pos;
        /// <summary>
        /// Lin FID/PID
        /// </summary>
        public int ID;
    }
    public struct __CanOutMessage
    {
        /// <summary>
        /// 데이타
        /// </summary>
        public byte[] Data;
        /// <summary>
        /// Lin FID/PID
        /// </summary>
        public int ID;
    }
    public struct __CanOutPos
    {
        /// <summary>
        /// 데이타
        /// </summary>
        public byte Data;
        /// <summary>
        /// 초기화 데이타 
        /// </summary>
        public byte Mask;
        //시작 위치
        public short Byte;
        /// <summary>
        /// 비트 위치
        /// </summary>
        public short Pos;
        /// <summary>
        /// Lin FID/PID
        /// </summary>
        public int ID;
    }

    public struct PSeatPowrItem
    {
        public int Batt;
        public int Gnd;
    }

    public struct PSeatPower
    {
        public PSeatPowrItem Batt1;
        public PSeatPowrItem Batt2;
    }
    public struct PinMapItem
    {
        public int PinNo;
        public int Mode; // 0 - B+ , 1 - Gnd
    }

    public class PSeatRuNMode
    {
        public const short Batt = 0;
        public const short Gnd = 1;
    }

    public struct PinMapStruct
    {
        //public PinMapItem IMSSet;
        //public PinMapItem IMSM1;
        //public PinMapItem IMSM2;
        public PinMapItem SlideFWD;
        public PinMapItem SlideBWD;
        public PinMapItem ReclineFWD;
        public PinMapItem ReclineBWD;
        public PinMapItem HeightUp;
        public PinMapItem HeightDn;
        public PinMapItem TiltUp;
        public PinMapItem TiltDn;

        public PSeatPowrItem LumberFwdBwd;
        public PSeatPowrItem LumberUpDn;

        public bool PSeat_직구동;
        public bool WalkIn;
        public PSeatPower Power;
    }
    public enum LHD_RHD
    {
        LHD,
        RHD
    }
    public enum PSEAT_TYPE
    {
        IMS,
        POWER,
        MANUAL
    }
    public struct __CheckItem__
    {
        public PSEAT_TYPE PSeat;
        //public bool Vent;
        //public bool Heater;
        //public bool Tilt;
        //public bool Leather; //가죽사양
        //public bool Cloth; //천 사양
        //public bool Lumber2Way;
        //public bool Lumber4Way;

        public bool Slide;
        public bool Tilt;
        public bool Height;
        public bool Recline;

        public bool WalkIn;
        public bool PSeat12Way;
        public bool PSeat10Way;
        public bool PSeat8Way;
        public bool PSeat4Way;
        //public bool PSeat2Way;            

        public bool Can;
        public bool ProductTestRunFlag;
        public LH_RH LhRh;
        public LHD_RHD LhdRhd;
        public bool Sound;
    }
    public struct __SpecItem__
    {
        public __MinMax__ IMS;
        public __MinMax__ Slide;
        public __MinMax__ Tilt;
        public __MinMax__ Height;
        public __MinMax__ Recliner;
        public __MinMax__ LumberUpDn;
        public __MinMax__ LumberFwdBwd;
    }
    public struct __SpecItem2__
    {
        public __MinMax__ SlideFwd;
        public __MinMax__ SlideBwd;
        public __MinMax__ TiltUp;
        public __MinMax__ TiltDn;
        public __MinMax__ HeightUp;
        public __MinMax__ HeightDn;
        public __MinMax__ ReclinerFwd;
        public __MinMax__ ReclinerBwd;
    }
    public struct __SpecItem3__
    {
        public float Slide;
        public float Tilt;
        public float Height;
        public float Recliner;
    }

    public struct __DeliveryPos__
    {
        public short Slide;
        public short Tilt;
        public short Height;
        public short Recliner;
    }

    public struct __SoundSpec__
    {
        public float StartMax;
        public float RunMax;
        /// <summary>
        /// 소음 측정시 Max 범위를 벗어났을 경우 그 소음이 SoundCheckRange 이사 유지 되면 인식한다.
        /// </summary>
        public float StartTime;
        public float 기동음구동음구분사이시간;
        public bool RMSMode;
        public float Offset;
    }

    public struct __Spec__
    {
        public string ModelName;
        public PinMapStruct PinMap;

        public __SpecItem__ Current;
        public __SpecItem2__ MovingSpeed;
        public __SpecItem3__ MovingStroke;
        public __SoundSpec__ Sound;
        public __DeliveryPos__ DeliveryPos;
        //-----------------------------
        public float SlideLimitTime;
        public float TiltLimitTime;
        public float HeightLimitTime;
        public float ReclinerLimitTime;
        public float LumberLimitTime;
        //-----------------------------
        public float SlideTestTime;
        public float TiltTestTime;
        public float HeightTestTime;
        public float ReclinerTestTime;
        public float LumberFwdBwdTestTime;
        public float LumberUpDnTestTime;
        //-----------------------------
        //public float SoundCheckTimeRange;
        public float SlideSoundCheckTime;
        public float TiltSoundCheckTime;
        public float HeightSoundCheckTime;
        public float ReclinerSoundCheckTime;
        public float LumberSoundCheckTime;        
        //-----------------------------
        public float SlideLimitCurr;
        public float TiltLimitCurr;
        public float HeightLimitCurr;
        public float ReclinerLimitCurr;

        public float TestVolt;
        //public bool Can;
    }

    public class SoundLevePos
    {
        public const short SLIDE_FWD = 0;
        public const short SLIDE_BWD = 1;
        public const short RECLINE_FWD = 2;
        public const short RECLINE_BWD = 3;
        public const short HEIGHT_UP = 4;
        public const short HEIGHT_DN = 5;
        public const short TILT_UP = 6;
        public const short TILT_DN = 7;
        public const short LUMBER_UP = 8;
        public const short LUMBER_DN = 9;
        public const short LUMBER_FWD = 10;
        public const short LUMBER_BWD = 11;        
    }
}