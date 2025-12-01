package excel

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	log "github.com/sirupsen/logrus"

	"github.com/xuri/excelize/v2"
)

type Point struct {
	Seq          string `excel:"序号"`
	BusinessUnit string `excel:"事业部"`
	Line         string `excel:"产线"`
	Area         string `excel:"区域"`
	Equipment    string `excel:"设备"`
	SubEquipment string `excel:"分部设备"`
	PointName    string `excel:"点位名称"`
	SensorType   string `excel:"传感器类型"`
	DataType     string `excel:"数据类型"`
	Precision    string `excel:"精度"`
	Range        string `excel:"取值范围"`

	Frequency      time.Duration `excel:"采集频率"` // 自动解析 + 默认补 ms
	Unit           string        `excel:"数据单位"`
	DataSourceAddr string        `excel:"数据源地址"`
	IOAddr         string        `excel:"IO地址"`
	DeviceCode     string        `excel:"设备编号"`
	DeviceSubCode  string        `excel:"设备附属编号"`
	PointCode      uint64        `excel:"点位编号"`
	PointExtraCode string        `excel:"点位额外编号"`
	GroupID        string        `excel:"分组编号"`
	NeedStore      string        `excel:"是否存储"`
	NeedPublish    string        `excel:"是否推送"`
	CalcType       string        `excel:"计算类型"`
	PublishTopic   string        `excel:"推送主题"`

	SheetName string `excel:"-"`
	RowNumber int    `excel:"-"`
}

func Load() {
	f, err := excelize.OpenFile("/Users/xiezg/Downloads/采集点位(34).xlsx" )
	if err != nil {
		log.Fatal("打开文件失败:", err)
	}

	var allPoints []Point

	for _, sheetName := range f.GetSheetList() {
		fmt.Printf("\n=== 正在解析 Sheet: %s ===\n", sheetName)
		rows, err := f.GetRows(sheetName)
		if err != nil || len(rows) < 2 {
			continue
		}

		header := rows[0]
		colIndex := make(map[string]int)
		for i, h := range header {
			colIndex[strings.TrimSpace(h)] = i
		}

		for rowIdx, row := range rows[1:] {
			if len(row) == 0 {
				continue
			}

			var p Point
			p.SheetName = sheetName
			p.RowNumber = rowIdx + 2

			v := reflect.ValueOf(&p).Elem()
			t := v.Type()

			errorMsgs := []string{}

			for i := 0; i < t.NumField(); i++ {
				field := t.Field(i)
				colName := field.Tag.Get("excel")
				if colName == "" || colName == "-" {
					continue
				}

				idx, ok := colIndex[colName]
				if !ok || idx >= len(row) {
					continue
				}

				cellValue := strings.TrimSpace(row[idx])

				switch field.Name {
				case "PointCode":
					if cellValue == "" {
						errorMsgs = append(errorMsgs, "点位编号为空")
						continue
					}
					if val, err := strconv.ParseUint(cellValue, 10, 64); err != nil {
						errorMsgs = append(errorMsgs, fmt.Sprintf("点位编号非法: %v", err))
					} else {
						v.Field(i).SetUint(val)
					}

				case "Frequency":
					if cellValue == "" {
						errorMsgs = append(errorMsgs, "采集频率为空")
						continue
					}
					parseStr := cellValue
					// 关键：没单位自动补 ms！
					if !strings.ContainsAny(parseStr, "nsuµmh") && 
						!strings.HasSuffix(strings.ToLower(parseStr), "ms") &&
						!strings.HasSuffix(strings.ToLower(parseStr), "s") {
						parseStr += "ms"
					}
					if dur, err := time.ParseDuration(parseStr); err != nil {
						errorMsgs = append(errorMsgs, fmt.Sprintf("采集频率解析失败: %v → %s", err, cellValue))
					} else if dur <= 0 {
						errorMsgs = append(errorMsgs, "采集频率必须大于0")
					} else {
						v.Field(i).Set(reflect.ValueOf(dur))
					}

				default:
					if field.Type.Kind() == reflect.String {
						v.Field(i).SetString(cellValue)
					}
				}
			}

			if len(errorMsgs) > 0 {
				log.Printf("Sheet[%s] 第 %d 行 错误: %v", sheetName, p.RowNumber, errorMsgs)
				continue
			}

			allPoints = append(allPoints, p)
		}
	}

	// 最终结果
	fmt.Printf("\n成功解析 %d 条有效点位\n\n", len(allPoints))
	for i, p := range allPoints {
		fmt.Printf("[%4d] %s | 编号=%d | 频率=%v | 存储=%s 推送=%s\n",
			i+1,
			p.PointName,
			p.PointCode,
			p.Frequency, // 自动打印 500ms、1s 等
			p.NeedStore,
			p.NeedPublish,
		)
	}
}
