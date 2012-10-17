using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CsCheck
{
    class ErrCode
    {
        public static string errMsg(int nErrCode)
        {
            string errMsg = "";
            switch (nErrCode)
            {
                case 4000:
                    errMsg = "讀卡機未連線";
                    break;
                case 4013:
                    errMsg = "未置入健保IC卡";
                    break;
                case 4014:
                    errMsg = "未置入醫事人員卡";
                    break;
                case 4029:
                    errMsg = "IC卡權限不足";
                    break;
                case 4033:
                    errMsg = "所置入非健保IC卡";
                    break;
                case 4034:
                    errMsg = "所置入非醫事人員卡";
                    break;
                case 4042:
                    errMsg = "醫事人員卡PIN尚未認證成功";
                    break;
                case 4050:
                    errMsg = "安全模組尚未與IDC認證";
                    break;
                case 4061:
                    errMsg = "網路不通";
                    break;
                case 4071:
                    errMsg = "健保IC卡與IDC認證失敗";
                    break;
                case 5001:
                    errMsg = "就醫可用次數不足";
                    break;
                case 5002:
                    errMsg = "卡片已註銷";
                    break;
                case 5003:
                    errMsg = "卡片已過有限期限";
                    break;
                case 5004:
                    errMsg = "新生兒依附就醫已逾60日";
                    break;
                case 5005:
                    errMsg = "讀卡機的就診日期時間讀取失敗";
                    break;
                case 5006:
                    errMsg = "讀取安全模組內的「醫療院所代碼」失敗";
                    break;
                case 5007:
                    errMsg = "寫入一組新的「就醫資料登錄」失敗";
                    break;
                case 5008:
                    errMsg = "安全模組簽章失敗";
                    break;
                case 5009:
                    errMsg = "同一天看診兩科(含)以上";
                    break;
                case 5012:
                    errMsg = "此人未在保。";
                    break;
                case 5015:
                    errMsg = "HC卡「門診處方箋」讀取失敗。";
                    break;
                case 5016:
                    errMsg = "HC卡「長期處方箋」讀取失敗。";
                    break;
                case 5017:
                    errMsg = "HC卡「重要醫令」讀取失敗。";
                    break;
                case 5018:
                    errMsg = "HC卡「過敏藥物」讀取失敗。";
                    break;
                case 5020:
                    errMsg = "要寫入的資料和健保IC卡不是屬於同一人。";
                    break;
                case 5022:
                    errMsg = "找不到「就醫資料登錄」中的該組資料。";
                    break;
                case 5023:
                    errMsg = "「就醫資料登錄」寫入失敗。";
                    break;
                case 5056:
                    errMsg = "讀取醫事人員ID失敗。";
                    break;
                case 5081:
                    errMsg = "最近24小時內同院所未曾就醫，故不可取消就醫。";
                    break;
                case 5102:
                    errMsg = "使用者所輸入之pin 值，與卡上之pin值不合。";
                    break;
                case 5109:
                    errMsg = "密碼輸入過程按『取消』鍵";
                    break;
                case 9129:
                    errMsg = "持卡人於非所限制的醫療院所就診";
                    break;
                case 9130:
                    errMsg = "醫事人員卡已失效";
                    break;
                case 9140:
                    errMsg = "醫事人員卡已逾有效期限";
                    break;
                default:
                    errMsg = "錯誤代碼：" + nErrCode;
                    break;
            }
            return errMsg;
        }
    }
}
