package utils

import (
	log "github.com/kdar/factorlog"

	"os"
	"strings"
)

func SetGlobalLogger(logPath string, level string) *log.FactorLog {
	if level == "" {
		level = "info"
	}
	sfmt := `%{Color "red:white" "CRITICAL"}%{Color "red" "ERROR"}%{Color "yellow" "WARN"}%{Color "green" "INFO"}%{Color "cyan" "DEBUG"}%{Color "blue" "TRACE"}[%{Date} %{Time}] [%{SEVERITY}:%{ShortFile}:%{Line}] %{Message}%{Color "reset"}`
	logger := log.New(os.Stdout, log.NewStdFormatter(sfmt))
	if len(logPath) > 0 {
		logf, err := os.OpenFile(logPath, os.O_WRONLY|os.O_APPEND|os.O_CREATE, 0640)
		if err != nil {
			return logger
		}
		logger = log.New(logf, log.NewStdFormatter(sfmt))
	}
	switch strings.ToLower(level) {
	case "debug":
		logger.SetSeverities(log.DEBUG | log.INFO | log.WARN | log.ERROR | log.FATAL | log.CRITICAL)
	case "info":
		logger.SetSeverities(log.INFO | log.WARN | log.ERROR | log.FATAL | log.CRITICAL)
	case "warn":
		logger.SetSeverities(log.WARN | log.ERROR | log.FATAL | log.CRITICAL)
	case "error":
		logger.SetSeverities(log.ERROR | log.FATAL | log.CRITICAL)
	default:
		logger.SetSeverities(log.INFO | log.WARN | log.ERROR | log.FATAL | log.CRITICAL)
	}

	return logger
}
