package schedule

const (
	RANK_DAY   uint8 = 1
	RANK_NIGHT uint8 = 2
)

const (
	WORK_SHOP_YANGZHA = "杨闸"
	WORK_SHOP_NANCE   = "南侧"
	WORK_SHOP_JIAHUI  = "佳汇"
)

type Schedule struct {
	StaffName    string `json:"staff_name" codec:"staff_name,omitempty"`       // 员工姓名
	ScheduleDate string `json:"schedule_date" codec:"schedule_date,omitempty"` // 考勤日期，格式：YYYY-MM-DD
	StartTime    string `json:"start_time" codec:"start_time,omitempty"`       // 排班开始时间，格式：YYYY-MM-DD HH24:MI
	EndTime      string `json:"end_time" codec:"end_time,omitempty"`           // 排班结束时间，格式：YYYY-MM-DD HH24:MI
	Rank         uint8  `json:"rank" codec:"rank,omitempty"`                   // 班次，1: 白班, 6:00-24:00, 2:夜班, 20:00-10:00
	WorkShop     string `json:"work_shop" codec:"work_shop,omitempty"`         // 工作店铺

	WageDayHours   float64 `json:"wage_day_hours" codec:"wage_day_hours,omitempty"`     // 结算白班小时数
	WageNightHours float64 `json:"wage_night_hours" codec:"wage_night_hours,omitempty"` // 结算夜班小时数
}

type SchedulesData struct {
	Data map[string][]*Schedule

	StartDate string
	EndDate   string
}

type HourStatData struct {
	Data map[string]map[string]int // work_shop:start-hour_end-hour:cnt

	StartDate string
	EndDate   string
}
