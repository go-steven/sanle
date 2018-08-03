package utils

import (
	"time"
)

func GetMonthDays(year, month int) (days int) {
	switch month {
	case 1, 3, 5, 7, 8, 10, 12:
		days = 31
	case 4, 6, 9, 11:
		days = 30
	case 2:
		if year%4 == 0 && (year%100 != 0 || year%400 == 0) {
			days = 29
		} else {
			days = 28
		}
	}
	return
}

func GetWeekDay(t time.Time) string {
	weekDays := map[time.Weekday]string{
		0: "日",
		1: "一",
		2: "二",
		3: "三",
		4: "四",
		5: "五",
		6: "六",
	}

	return weekDays[t.Weekday()]
}
