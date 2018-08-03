package kaoqing

import (
	"time"

	"github.com/go-steven/sanle/schedule"
)

func AggregateByStaff(kaoqings map[string][]*KaoQing, startDate, endDate string) (map[string]*KaoQingAggregate, error) {
	ret := make(map[string]*KaoQingAggregate)

	today := time.Now().Format("2006-01-02")

	for staffName, records := range kaoqings {
		aggregate := &KaoQingAggregate{
			StaffName: staffName,
		}

		for _, v := range records {
			if startDate != "" && v.ScheduleDate < startDate {
				continue
			}
			if endDate != "" && v.ScheduleDate > endDate {
				continue
			}

			if v.ScheduleDate >= today {
				continue
			}

			if v.EarlyMinutes > 0 {
				aggregate.EarlyDays++
			}
			if v.LateMinutes > 0 {
				aggregate.LateDays++
			}
			if v.OverMinutes > 0 {
				aggregate.OverDays++
			}
			switch v.RecordSts {
			case RECORD_STS_NONE:
				aggregate.RecordNoneDays++
			case RECORD_STS_NOSTART:
				aggregate.RecordNoStartDays++
			case RECORD_STS_NOEND:
				aggregate.RecordNoEndDays++
			case RECORD_STS_NORMAL:
				aggregate.RecordNormalDays++
			}

			if v.Sts == KAOQING_STS_NORMAL {
				aggregate.NormalDays++
			}

			switch v.WorkShop {
			case schedule.WORK_SHOP_YANGZHA:
				aggregate.YangzhaDays++
			case schedule.WORK_SHOP_NANCE:
				aggregate.NanceDays++
			case schedule.WORK_SHOP_JIAHUI:
				aggregate.JiahuiDays++
			}
			aggregate.TotalDays++
		}

		ret[staffName] = aggregate
	}
	return ret, nil
}
