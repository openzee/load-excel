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

	Frequency       time.Duration `excel:"采集频率"` // 自动解析 + 默认补 ms
	Unit            string        `excel:"数据单位"`
	DataSourceAddr  string        `excel:"数据源地址"`
	IOAddr          string        `excel:"IO地址"`
	DeviceCode      string        `excel:"设备编号"`
	DeviceSubCode   string        `excel:"设备附属编号"`
	PointPrimaryKey uint64        `excel:"点位编号"`
	PointExtraCode  string        `excel:"点位额外编号"`
	GroupID         string        `excel:"分组编号"`
	NeedStore       string        `excel:"是否存储"`
	//NeedPublish    string        `excel:"是否推送"`
	//CalcType       string        `excel:"计算类型"`
	//PublishTopic   string        `excel:"推送主题"`

	SheetName string `excel:"-"`
	RowNumber int    `excel:"-"`
}

func (obj Point) String() string {
	return fmt.Sprintf("SheetName:%v RowNumber:%v 点位编号:%v 数据源地址:%v IO地址:%v", 
		obj.SheetName,
		obj.RowNumber,
		obj.PointPrimaryKey,
		obj.DataSourceAddr,
		obj.IOAddr,
	)
}

// 将相同地址和频率的点位重新组装
func ReassembleWithAddrAndFreq(pt []Point) map[string]map[time.Duration][]Point {

	mapWithDataAddr := make(map[string]map[time.Duration][]Point)
	count := 0

	for _, p := range pt {

		mapWithFreq, ok := mapWithDataAddr[p.DataSourceAddr]
		if !ok {
			mapWithFreq = make(map[time.Duration][]Point)
			mapWithDataAddr[p.DataSourceAddr] = mapWithFreq
		}

		pts, ok := mapWithFreq[p.Frequency]
		if !ok {
			pts = make([]Point, 0)
			mapWithFreq[p.Frequency] = pts
		}

		pts = append(pts, p)
		count = count + 1
		mapWithFreq[p.Frequency] = pts //pts可能已经被更新，需要回写到map中
	}

	if count != len(pt) {
		log.Fatalf("点位重组后，长度缺失")
	}

	return mapWithDataAddr
}

func ParseExcel(fname string, onlySheets ...string) ([]Point, error) {
	f, err := excelize.OpenFile(fname)
	if err != nil {
		return nil, fmt.Errorf("open file %v fails, err:%v", fname, err)
	}

	var allPoints []Point

	sheetList := f.GetSheetList()

	// 如果指定了 OnlySheets，则过滤
	if len(onlySheets) > 0 {
		nameMap := make(map[string]bool)
		for _, name := range onlySheets {
			nameMap[strings.TrimSpace(name)] = true
		}

		filtered := []string{}
		for _, name := range sheetList {

			//如果nameMap中没有，返回的值是false
			if nameMap[name] || nameMap[strings.TrimSpace(name)] {
				filtered = append(filtered, name)
			}
		}
		sheetList = filtered

		if len(sheetList) == 0 {
			return nil, fmt.Errorf("未找到指定的 Sheet: %v", onlySheets)
		}
	}

	// 遍历要解析的 Sheet
	for _, sheetName := range sheetList {
		log.Debugf("\n=== 正在解析 Sheet: %s ===\n", sheetName)
		rows, err := f.GetRows(sheetName)
		if err != nil || len(rows) < 2 {
			log.Errorf("sheet:%v load fails, rows:%v err:%v", sheetName, rows, err)
			continue
		}

		header := rows[0]
		colIndex := make(map[string]int)
		for i, h := range header {
			colIndex[strings.TrimSpace(h)] = i
		}

		for rowIdx, row := range rows[1:] {
			if len(row) == 0 {
				log.Errorf("sheet:%v rows:%v empty", sheetName, rowIdx+1)
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
					log.Warningf("sheetName:[%v] row:[%v] colName:[%v] idx:[%v] 未找到[%v]或超过row范围%v", sheetName, rowIdx+1, colName, idx, !ok, row )         
					continue
				}

				cellValue := strings.TrimSpace(row[idx])

				switch field.Name {
				case "PointPrimaryKey":
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

	return allPoints, nil
}
