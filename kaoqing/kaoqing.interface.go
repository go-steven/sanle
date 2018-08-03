package kaoqing

import (
	"github.com/go-steven/sanle/schedule"
)

type OrigRecord struct {
	StaffName  string `json:"staff_name" codec:"staff_name,omitempty"`
	RecordDate string `json:"record_date" codec:"record_date,omitempty"`
	RecordOn   string `json:"record_on" codec:"record_on,omitempty"`
	WorkShop   string `json:"work_shop" codec:"work_shop,omitempty"`
}

const (
	RECORD_STS_NORMAL  = 1
	RECORD_STS_NONE    = 2
	RECORD_STS_NOSTART = 3
	RECORD_STS_NOEND   = 4
)

const (
	KAOQING_STS_NORMAL         = 1
	KAOQING_STS_LATE           = 2
	KAOQING_STS_EARLY          = 3
	KAOQING_STS_LATE_AND_EARLY = 4
	KAOQING_STS_NONE           = 9
)

type KaoQing struct {
	schedule.Schedule

	StartRecord  *OrigRecord
	EndRecord    *OrigRecord
	RecordSts    uint8 `json:"record_sts" codec:"record_sts,omitempty"`       // 打卡状态, 1-正常, 2-无打卡记录, 3-上班未打卡, 4-下班未打卡
	Sts          uint8 `json:"sts" codec:"sts,omitempty"`                     // 考勤状态, 1-正常, 2-迟到, 3-早退, 4-迟到+早退， 5，未打卡或异常
	LateMinutes  int   `json:"late_minutes" codec:"late_minutes,omitempty"`   // 迟到分钟数
	EarlyMinutes int   `json:"early_minutes" codec:"early_minutes,omitempty"` // 早退分钟数
	OverMinutes  int   `json:"over_minutes" codec:"over_minutes,omitempty"`   // 加班分钟数
}

type KaoQingAggregate struct {
	StaffName         string `json:"staff_name" codec:"staff_name,omitempty"`                   // 员工姓名
	TotalDays         int    `json:"total_days" codec:"total_days,omitempty"`                   // 总上班天数
	YangzhaDays       int    `json:"yangzha_days" codec:"yangzha_days,omitempty"`               // 杨闸天数
	NanceDays         int    `json:"nance_days" codec:"nance_days,omitempty"`                   // 南侧天数
	JiahuiDays        int    `json:"jiahui_days" codec:"jiahui_days,omitempty"`                 // 佳汇天数
	NormalDays        int    `json:"normal_days" codec:"normal_days,omitempty"`                 // 正常天数
	LateDays          int    `json:"late_days" codec:"late_days,omitempty"`                     // 迟到天数
	EarlyDays         int    `json:"early_days" codec:"early_days,omitempty"`                   // 早退天数
	OverDays          int    `json:"over_days" codec:"over_days,omitempty"`                     // 加班天数
	RecordNormalDays  int    `json:"record_normal_days" codec:"record_normal_days,omitempty"`   // 打卡正常天数
	RecordNoneDays    int    `json:"record_none_days" codec:"record_none_days,omitempty"`       // 全部未打卡天数
	RecordNoStartDays int    `json:"record_nostart_days" codec:"record_nostart_days,omitempty"` // 上班未打卡天数
	RecordNoEndDays   int    `json:"record_noend_days" codec:"record_noend_days,omitempty"`     // 下班未打卡天数
}
