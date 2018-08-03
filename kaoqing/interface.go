package kaoqing

import (
	log "github.com/kdar/factorlog"

	"github.com/go-steven/sanle/schedule"
	"github.com/go-steven/sanle/utils"
)

var Logger *log.FactorLog = utils.SetGlobalLogger("", "info")

// SetLogger 初始化设置Logger
func SetLogger(alogger *log.FactorLog) {
	Logger = alogger
	schedule.SetLogger(alogger)
}
