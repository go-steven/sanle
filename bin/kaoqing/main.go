package main

import (
	log "github.com/kdar/factorlog"

	"flag"
	"fmt"
	"math/rand"
	"time"

	"github.com/go-steven/sanle/kaoqing"
	"github.com/go-steven/sanle/schedule"
	"github.com/go-steven/sanle/utils"
)

var (
	logFlag           = flag.String("log", "", "set log path")
	scheduleExcelFlag = flag.String("schedule-excel", "d:/排班表.xlsx", "input staff schedule excel.")
	recordExcelFlag   = flag.String("record-excel", "C:/Users/m1358/Downloads/杨闸南侧佳汇大群_考勤报表_20180726-20180803.xlsx", "input staff kaoqing orig record excel.")
	startDateFlag     = flag.String("start-date", "2018-07-26", "input start date.")
	endDateFlag       = flag.String("end-date", "2018-08-25", "input end date.")

	Logger *log.FactorLog
)

func init() {
	rand.Seed(time.Now().UnixNano())
}

func main() {
	flag.Parse()
	Logger = utils.SetGlobalLogger(*logFlag, "debug")
	kaoqing.SetLogger(Logger)
	schedule.SetLogger(Logger)
	startDate := *startDateFlag
	endDate := *endDateFlag

	if endDate >= time.Now().Format("2006-01-02") {
		endDate = time.Now().Add((-1) * 24 * time.Hour).Format("2006-01-02") // 取昨天
	}

	schedulesData, err := schedule.ReadScheduleFromExcel(*scheduleExcelFlag, []string{
		schedule.WORK_SHOP_YANGZHA,
		schedule.WORK_SHOP_NANCE,
		schedule.WORK_SHOP_JIAHUI,
	})
	if err != nil {
		Logger.Error(err)
		return
	}
	Logger.Infof("schedulesData: %s", utils.Json(schedulesData))

	origRecords, err := kaoqing.ReadOrignRecordFromExcel(*recordExcelFlag, startDate, endDate)
	if err != nil {
		Logger.Error(err)
		return
	}

	kaoqings, err := kaoqing.UpdateUserKaoQing(schedulesData.Data, origRecords, startDate, endDate)
	if err != nil {
		Logger.Error(err)
		return
	}

	kaoqingAggregate, err := kaoqing.AggregateByStaff(kaoqings, startDate, endDate)
	if err != nil {
		Logger.Error(err)
		return
	}

	filterName := ""
	if err := kaoqing.SaveToExcel(schedule.GenerateOutputFilename(*scheduleExcelFlag, fmt.Sprintf("%s_%s_%s%s", "考勤记录", startDate, endDate, filterName)), origRecords, kaoqings, kaoqingAggregate, startDate, endDate, filterName); err != nil {
		Logger.Error(err)
		return
	}
}
