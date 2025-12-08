using System;
using System.Text;
using System.Runtime.InteropServices;

namespace EtherCAT_DLL_Err
{
    public class CEtherCAT_DLL_Err
    {
        public const ushort ERR_ECAT_NO_ERROR					    =   0x0;
        public const ushort ERR_ECAT_HW_NO_INITIALIZE			    =   0x1;
        public const ushort ERR_ECAT_HW_PWM_INITIAL					=   0x2;
        public const ushort ERR_ECAT_HW_HAS_INITIALIZED				=   0x3;

        public const ushort ERR_ECAT_EEPROM_READ				    =   0x10;
        public const ushort ERR_ECAT_EEPROM_WRITE				    =   0x11;
        public const ushort ERR_ECAT_ENVIRONMENT_RECORD_DISABLE		=   0x12;
        public const ushort ERR_ECAT_ENVIRONMENT_RECORD_NO_MATCH	=   0x13;
        public const ushort ERR_ECAT_ENVIRONMENT_RECORD_FILE_OPEN 	=   0x14;
        public const ushort ERR_ECAT_ENVIRONMENT_RECORD_NOT_CREATE	=   0x15;
        public const ushort ERR_ECAT_XML_FILE_PATH					=   0x16;
        public const ushort ERR_ECAT_DEVICE_OPEN					=   0x17;
        public const ushort ERR_ECAT_NO_DEVICE						=   0x18;
        public const ushort ERR_ECAT_NO_MASTER						=   0x19;
        public const ushort ERR_ECAT_NO_SLAVE						=   0x1A;
        public const ushort ERR_ECAT_UNKNOWN_SLAVE					=   0x1B;
        public const ushort ERR_ECAT_IST_CREATE						=   0x1C;
        public const ushort ERR_ECAT_MASTER_CREATE					=   0x1D;
        public const ushort ERR_ECAT_MASTER_REQUEST_STATE			=   0x1E;
        public const ushort ERR_ECAT_MASTER_OPERATION_NOT_READY		=   0x1F;
        public const ushort ERR_ECAT_DELTA_NODE_ID_ALIAS_READ		=   0x20;

        public const ushort ERR_ECAT_PIPELINE_CORE_TIMER_CREATE		=   0x80;
        public const ushort ERR_ECAT_PIPELINE_CREATE			    =   0x81;
        public const ushort ERR_ECAT_COMMAND_ENQUEUE			    =   0x82;
        public const ushort ERR_ECAT_API_BUFFER_ENQUEUE				=   0x83;

        public const ushort ERR_ECAT_NODE_ID						=   0x100;
        public const ushort ERR_ECAT_SLOT_ID						=   0x101;
        public const ushort ERR_ECAT_SDO_DOWNLOAD				    =   0x102;
        public const ushort ERR_ECAT_SDO_UPLOAD						=   0x103;
        public const ushort ERR_ECAT_GET_PROCESS_DATA				=   0x104;

        public const ushort ERR_ECAT_DIO_CHANNEL_INVALID			=   0x200;
        public const ushort ERR_ECAT_ADDA_CHANNEL_INVALID		    =   0x201;
        public const ushort ERR_ECAT_MOTION_NOT_FINISHED			=   0x202;
        public const ushort ERR_ECAT_SET_PULSE_MODE					=   0x203;
        public const ushort ERR_ECAT_SET_SOFTLIMIT					=   0x204;
        public const ushort ERR_ECAT_SET_POSITION					=   0x205;
        public const ushort ERR_ECAT_GET_SPEED						=   0x206;
        public const ushort ERR_ECAT_GET_MCDONE						=   0x207;
        public const ushort ERR_ECAT_SET_HOME_CONFIG				=   0x208;
        public const ushort ERR_ECAT_SET_P2P_CONFIG					=   0x209;
        public const ushort ERR_ECAT_SET_PT_CONFIG					=   0x20A;
        public const ushort ERR_ECAT_SET_PV_CONFIG					=   0x20B;
        public const ushort ERR_ECAT_SET_CSP_CONFIG					=   0x20C;
        public const ushort ERR_ECAT_SET_MULTI_AXES_LINE_CONFIG		=   0x20D;
        public const ushort ERR_ECAT_SET_MULTI_AXES_ARC_CONFIG		=   0x20E;

        public const ushort ERR_ECAT_MD1_SET_GEAR				    =   0x300;
        public const ushort ERR_ECAT_MD1_SET_P_CHANGE			    =   0x301;
        public const ushort ERR_ECAT_MD1_SET_V_CHANGE			    =   0x302;
        public const ushort ERR_ECAT_MD1_SET_SOFTLIMIT				=   0x303;
        public const ushort ERR_ECAT_MD1_SET_SLD				    =   0x304;
        public const ushort ERR_ECAT_MD1_SET_HOME_CONFIG		    =   0x305;
        public const ushort ERR_ECAT_MD1_SET_P2P_CONFIG				=   0x306;
        public const ushort ERR_ECAT_MD1_SET_V_MOVE_CONFIG			=   0x307;
        public const ushort ERR_ECAT_MD1_SET_LINE_CONFIG		    =   0x308;
        public const ushort ERR_ECAT_MD1_SET_ARC_CONFIG				=   0x309;

        public const ushort ERR_ECAT_PATH_NOT_SUPPORT				=   0x400;
        public const ushort ERR_ECAT_PATH_AXIS_NUM					=   0x401;
        public const ushort ERR_ECAT_PATH_AXIS_NO					=   0x402;
        public const ushort ERR_ECAT_PATH_PARA						=   0x403;
        public const ushort ERR_ECAT_PATH_ISR_FUNC_EVENT		    =   0x404;
        public const ushort ERR_ECAT_PATH_AXISNO_UNDER_GROUP		=   0x405;

        public const ushort ERR_ECAT_PATH_ROBOT_NOT_SUPPORT			=   0x480;
        public const ushort ERR_ECAT_PATH_ROBOT_STOP				=   0x481;
        public const ushort ERR_ECAT_PATH_ROBOT_AXIS_OVERFLOW	    =   0x482;
        public const ushort ERR_ECAT_PATH_ROBOT_BUFFER_FULL			=   0x483;

        public const ushort ERR_ESI_INITIAL							=   0xF00;
        public const ushort ERR_ESI_OPEN_DEVICE						=   0xF01;
        public const ushort ERR_ESI_CREATE_CANOPEN_OD_LIST			=   0xF02;
        public const ushort ERR_ESI_NO_DATA_TYPE_INFO			    =   0xF03;
        public const ushort ERR_ESI_NO_OBJECT_INFO					=   0xF04;
        public const ushort ERR_ESI_CREATE_SYNC_MANAGER				=   0xF05;
        public const ushort ERR_ESI_CREATE_FMMU_CONTROL				=   0xF06;
        public const ushort ERR_ESI_NO_PDO_CHANNEL					=   0xF07;
        public const ushort ERR_ESI_NO_PDO_MAPPING					=   0xF08;
        public const ushort ERR_ESI_PDO_MAPPING_INSERT				=   0xF09;
        public const ushort ERR_ESI_PDO_MAPPING_DELETE				=   0xF0A;
        public const ushort ERR_ESI_CREATE_DISTRIBUTED_CLOCK		=   0xF0B;

        public const ushort ERR_ESI_ENI_INFORMATION_INITIAL			=   0xFF0;
        public const ushort ERR_ESI_ENI_FILE_INITIAL			    =   0xFF1;
        public const ushort ERR_ESI_ENI_FILE_SAVE				    =   0xFF2;

        public const ushort ERR_ECAT_NO_SLAVE_FOUND					=   0x1000;
        public const ushort ERR_ECAT_INITIAL_TIMEOUT				=   0x1001;
        public const ushort ERR_ECAT_MODE_CHANGE_FAILED				=   0x1002;
        public const ushort ERR_ECAT_SLAVE_ID					    =   0x1003;
        public const ushort ERR_ECAT_ALIAS_SLAVE_ID					=   0x1004;

        public const ushort ERR_ECAT_NEED_INITIAL				    =   0x1100;
        public const ushort ERR_ECAT_NEED_RESET						=   0x1101;
        public const ushort ERR_ECAT_NEED_CONNECT				    =   0x1102;
        public const ushort ERR_ECAT_NEED_DC_OP						=   0x1103;
        public const ushort ERR_ECAT_NEED_RALM						=   0x1104;
        public const ushort ERR_ECAT_NEED_SVON						=   0x1105;
        public const ushort ERR_ECAT_NEED_HOMECONFIG				=   0x1106;

        public const ushort ERR_ECAT_RING_BUFFER_FULL			    =   0x1200;
        public const ushort ERR_ECAT_API_PARAMETER					=   0x1201;
        public const ushort ERR_ECAT_SLAVE_TYPE						=   0x1202;
        public const ushort ERR_ECAT_TARGET_REACHED					=   0x1203;
        public const ushort ERR_ECAT_MODE_NOT_SUPPORT			    =   0x1204;
        public const ushort ERR_ECAT_MOTION_TYPE				    =   0x1205;
        public const ushort ERR_ECAT_PDO_NOT_MAPPING                =   0x1206;

        public const ushort ERR_ECAT_PDO_TX_FAILED					=   0x1300;
        public const ushort ERR_ECAT_SDO_TIMEOUT					=   0x1301;
        public const ushort ERR_ECAT_SDO_RETURN						=   0x1302;
        public const ushort ERR_ECAT_PDO_RX_FAILED					=   0x1303;
        public const ushort ERR_ECAT_MAILBOX						=   0x1304;
        
        public const ushort ERR_ECAT_GROUP_NUMBER				    =   0x1400;
        public const ushort ERR_ECAT_GROUP_ENABLE				    =   0x1401;
        public const ushort ERR_ECAT_GROUP_PAUSE					=   0x1402;
        public const ushort ERR_ECAT_GROUP_SLAVE				    =   0x1403;
        public const ushort ERR_ECAT_GROUP_MODE						=   0x1404;
        public const ushort ERR_ECAT_GROUP_ALREADY_USED				=   0x1405;
        public const ushort ERR_ECAT_GROUP_TYPE						=   0x1406;
        public const ushort ERR_ECAT_GROUP_SVON						=   0x1407;
        public const ushort ERR_ECAT_GROUP_ALM						=   0x1408;
        public const ushort ERR_ECAT_GROUP_DATA_BUFFER				=   0x1409;
        public const ushort ERR_ECAT_GROUP_TIMEOUT					=   0x140A;

        public const ushort ERR_ECAT_SERVO_PARA_EMPTY			    =   0x1500;
        public const ushort ERR_ECAT_SERVO_PARA_RO					=   0x1501;

        public const ushort ERR_ECAT_RECORD_NEED_DISABLE			=   0x1600;

        public const ushort ERR_PATH_BOARD_INIIT					=   0x8001;
        public const ushort ERR_PATH_BOARD_NUMBER					=   0x8002;
        public const ushort ERR_PATH_INITIAL_BOARD_NUMBER			=   0x8003;
        public const ushort ERR_PATH_BASE_ADDR_ERROR				=   0x8004;
        public const ushort ERR_PATH_BASE_ADDR_CONFLICT				=   0x8005;
        public const ushort ERR_PATH_DUPLICATE_BOARD_SETTING		=   0x8006;
        public const ushort ERR_PATH_DUPLICATE_IRQ_SETTING			=   0x8007;
        public const ushort ERR_PATH_ENC_NUMBER						=   0x8008;
        public const ushort ERR_PATH_MODULE_NUMBER					=   0x8009;
        public const ushort ERR_PATH_TIMER_VALUE					=   0x800A;
        public const ushort ERR_PATH_ENABLE							=   0x800B;
        public const ushort ERR_PATH_RANGE							=   0x800C;
        public const ushort ERR_PATH_MEMORY_ALLOC				    =   0x800D;
        public const ushort ERR_PATH_MOTION_BUSY				    =   0x800E;
        public const ushort ERR_PATH_MOTION_NO_START			    =   0x800F;
        public const ushort ERR_PATH_WRONG_SPEED				    =   0x8010;
        public const ushort ERR_PATH_WRONG_ACCTIME					=   0x8011;
        public const ushort ERR_PATH_IO_ALAM					    =   0x8012;
        public const ushort ERR_PATH_OPEN_FILE_FAILED			    =   0x8013;
        public const ushort ERR_PATH_MEMORY_ALLOCATE			    =   0x8014;
        public const ushort ERR_PATH_MEMORY_NOT_FREE			    =   0x8015;
        public const ushort ERR_PATH_OUTPUT_FILE_NOT_CLOSE			=   0x8016;
        public const ushort ERR_PATH_MOVE_AXIS_NOT_MATCH		    =   0x8017;
        public const ushort ERR_PATH_PITCH_ZERO						=   0x8018;
        public const ushort ERR_PATH_TIMEOUT					    =   0x8019;
        public const ushort ERR_PATH_PCI_BIOS_NOT_EXIST				=   0x801A;
        public const ushort ERR_PATH_BUFFER_FULL				    =   0x801B;
        public const ushort ERR_PATH_ERROR							=   0x801C;
        public const ushort ERR_PATH_REACH_SWLIMIT					=   0x801D;
        public const ushort ERR_PATH_NO_SUPPRT_MODE					=   0x801E;
        public const ushort ERR_PATH_AXIS_CORRELATION			    =   0x801F;
        public const ushort ERR_PATH_FEEDHOLD_SUPPROT			    =   0x8020;
        public const ushort ERR_PATH_SD_STOP_ON						=   0x8021;
        public const ushort ERR_PATH_VELOCITY_CHANGE_SUPER			=   0x8022;
        public const ushort ERR_PATH_COMMAND_SET				    =   0x8023;
        public const ushort ERR_PATH_SDO_MESSAGE_CHOKE				=   0x8024;
        public const ushort ERR_PATH_VELOCITY_CHANGE_BUFFER_FEEDHOLD=   0x8025;
        public const ushort ERR_PATH_VELOCITY_CHANGE_SYNC_MOVE		=   0x8026;
        public const ushort ERR_PATH_VELOCITY_CHANGE_SD_ON			=   0x8027;
        public const ushort ERR_PATH_POS_CHANGE_MODE				=   0x8028;
        public const ushort ERR_PATH_BUFFER_LENGTH					=   0x8029;
        public const ushort ERR_PATH_2SEG_DISTANCE					=   0x802A;
        public const ushort ERR_PATH_ARC_CENTER_POSITION			=   0x802B;
        public const ushort ERR_PATH_ARC_END_POSITION			    =   0x802C;
        public const ushort ERR_PATH_ARC_ANGLE_CALC					=   0x802D;
        public const ushort ERR_PATH_ARC_RADIUS_CALC				=   0x802E;
        public const ushort ERR_PATH_GEAR_SETTING				    =   0x802F;
        public const ushort ERR_PATH_CAM_TABLE						=   0x8030;
        public const ushort ERR_PATH_AXES_NUMBER					=   0x8031;
        public const ushort ERR_PATH_SPIRAL_END_POSITION			=   0x8032;
        public const ushort ERR_PATH_SPEED_MODE_SLAVE			    =   0x8033;
        public const ushort ERR_PATH_SPEED_MODE_SET_SLAVE		    =   0x8034;
        public const ushort ERR_PATH_VELOCITY_CHANGE				=   0x8035;
        public const ushort ERR_PATH_BACKLASH_STEP					=   0x8036;
        public const ushort ERR_PATH_BACKLASH_STATUS				=   0x8037;
        public const ushort ERR_PATH_DIST_OVER						=   0x8038;
       
        public const ushort ERR_RTX_RTSS_LOAD					    =   0xD000;
        public const ushort ERR_RTX_CONNECT_LINK_FAILED				=   0xD001;
        public const ushort ERR_RTX_EVENT_FAILED					=   0xD002;
        public const ushort ERR_RTX_CONNECT_FAILED					=   0xD003;
        public const ushort ERR_RTX_CONFIG_EDITED				    =   0xD004;
        public const ushort ERR_RTX_SECURITY_FAILED					=   0xD005;
        public const ushort ERR_RTX_COMMANDING						=   0xD006;
        
        public const ushort ERR_RTX_WIN32_SYSTEM_NOT_SUPPORT		=   0xD100;
        public const ushort ERR_RTX_CALLBACK_CLOSE					=   0xD101;
        public const ushort ERR_RTX_CALLBACK_FUNCTION			    =   0xD102;
        public const ushort ERR_RTX_CALLBACK_THREAD					=   0xD103;

        public const ushort ERR_ECAT_DLL_IS_USED					=   0xF000;
        public const ushort ERR_ECAT_NO_DLL_FOUND				    =   0xF001;
        public const ushort ERR_ECAT_NO_RTSS_DLL_FOUND				=   0xF002;
        public const ushort ERR_ECAT_NO_CARD_DLL_FOUND				=   0xF003;
        public const ushort ERR_ECAT_NO_ESI_DLL_FOUND               =   0xF004;
        public const ushort ERR_ECAT_SAME_CARD_NUMBER			    =   0xF005;
        public const ushort ERR_ECAT_CARDNO_ERROR				    =   0xF006;
        public const ushort ERR_ECAT_GET_DLL_PATH				    =   0xF007;
        public const ushort ERR_ECAT_GET_DLL_VERSION				=   0xF008;
        public const ushort ERR_ECAT_NOT_SUPPORT                    =   0xF009;
    }
}
