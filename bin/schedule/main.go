package main

import (
	log "github.com/kdar/factorlog"

	"flag"
	"math/rand"
	"time"

	"github.com/go-steven/sanle/schedule"
	"github.com/go-steven/sanle/utils"
)

var (
	logFlag           = flag.String("log", "", "set log path")
	scheduleExcelFlag = flag.String("schedule-excel", "d:/排班表.xlsx", "input user schedule excel.")
	wageFlag          = flag.Bool("wage", true, "generate wage data.")
	hourStatFlag      = flag.Bool("hour-stat", true, "generate hour stats.")

	Logger *log.FactorLog
)

func init() {
	rand.Seed(time.Now().UnixNano())
}

func main() {
	flag.Parse()
	Logger = utils.SetGlobalLogger(*logFlag, "")
	schedule.SetLogger(Logger)

	schedulesData, err := schedule.ReadScheduleFromExcel(*scheduleExcelFlag, []string{
		schedule.WORK_SHOP_YANGZHA,
		schedule.WORK_SHOP_NANCE,
		schedule.WORK_SHOP_JIAHUI,
	})
	if err != nil {
		Logger.Error(err)
		return
	}
	Logger.Infof("schedules: %s", utils.Json(schedulesData))

	if *wageFlag {
		if err := schedule.SaveToExcel(schedule.GenerateOutputFilename(*scheduleExcelFlag, "工时结算"), schedulesData); err != nil {
			Logger.Error(err)
			return
		}
	}

	if *hourStatFlag {
		hourStatsData, err := schedule.AggregateByHour(schedulesData)
		if err != nil {
			Logger.Error(err)
			return
		}
		if err := schedule.SaveHourStatsToExcel(schedule.GenerateOutputFilename(*scheduleExcelFlag, "工时统计"), hourStatsData); err != nil {
			Logger.Error(err)
			return
		}
	}
}
