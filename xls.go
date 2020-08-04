package xlsxutil

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/miaomiao3/xlsx" // i fork this repo to enable SetType
	"gopkg.in/yaml.v2"
	"reflect"
	"strconv"
	"strings"
	"time"
)

const (
	xlsTag                = "xls"
	inlineTag             = "inline"
	defaultFloatPrecision = 6
)

// dump data to file
// data: ptr of slice, slice element should be a struct
// sep: separator of csv line. be careful to avoid value conflict
func XlsDump(file *xlsx.File, sheetName string, data interface{}) error {
	value := reflect.ValueOf(data)
	if value.Kind() != reflect.Slice {
		return errors.New("do not support non slice data")
	}
	l := value.Len()
	if l == 0 {
		return errors.New("empty slice")
	}
	dataSlice := make([]interface{}, l)
	for i := 0; i < l; i++ {
		if value.Index(i).CanInterface() {

			if value.Index(i).IsNil() {
				return errors.New("slice contain nil element")
			}

			//if value.Index(i).CanAddr() {
			//	return errors.New("slice contains element that can't addr")
			//}
			dataSlice[i] = value.Index(i).Interface()
		} else {
			return errors.New("slice item CanInterface failed")
		}

	}
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return err

	}
	var optionMap map[string]*xlsOption
	for k, v := range dataSlice {
		if k == 0 {
			row := sheet.AddRow()
			optionMap, err = addHeaderRow(row, v)
			if err != nil {
				return err
			}
		}
		row := sheet.AddRow()
		err = addRow(row, v, optionMap)
		if err != nil {
			return err
		}
	}
	return err
}

type xlsOption struct {
	Inline    bool `yaml:"inline"`
	Precision int  `yaml:"precision"`
}

// add header
func addHeaderRow(row *xlsx.Row, v interface{}) (optionMap map[string]*xlsOption, err error) {
	if row == nil {
		return optionMap, errors.New("row empty")
	}
	dataValue := reflect.ValueOf(v)
	dataType := reflect.TypeOf(v)
	if dataType == nil {
		return optionMap, fmt.Errorf("RowHeadAddStruct get nil input")
	}

	if dataType.Kind() == reflect.Ptr {
		dataValue = dataValue.Elem()
		dataType = dataType.Elem()
	}

	if dataType.Kind() != reflect.Struct {
		err = errors.New("only support struct or struct pointer")
		return
	}

	optionMap = make(map[string]*xlsOption)

	num := dataType.NumField()
	for i := 0; i < num; i++ {
		tag := dataType.Field(i).Tag.Get(xlsTag)
		if len(tag) == 0 {
			continue
		}

		// get option
		option := getOptionFromTag(tag)

		field := dataType.Field(i)
		fieldName := field.Name

		if option != nil {
			optionMap[fieldName] = option
		}

		fieldValue := dataValue.FieldByName(fieldName)
		if !fieldValue.IsValid() {
			continue
		}

		if field.Type.Kind() == reflect.Struct && option.Inline {
			newOptionMap, err := addHeaderRow(row, fieldValue.Interface())
			if err != nil {
				return nil, err
			}
			for k, v := range newOptionMap {
				optionMap[k] = v
			}
		} else {
			cell := row.AddCell()
			cell.Value = strings.Split(tag, ",")[0]
			cell.SetType(xlsx.CellTypeString)
		}
	}
	return optionMap, nil
}


func addRow(row *xlsx.Row, v interface{}, optionMap map[string]*xlsOption) (err error) {
	if row == nil {
		return errors.New("row empty")
	}
	dataType := reflect.TypeOf(v)
	dataValue := reflect.ValueOf(v)

	if dataType.Kind() == reflect.Ptr {
		originType := reflect.ValueOf(v).Elem().Type()
		if originType.Kind() != reflect.Struct {
			err = errors.New("only support struct pointer")
			return
		}
		dataValue = dataValue.Elem()
		dataType = dataType.Elem()
	} else if dataType.Kind() != reflect.Struct {
		err = errors.New("only support struct or struct pointer")
		return
	}

	num := dataType.NumField()
	for i := 0; i < num; i++ {
		tagStr := dataType.Field(i).Tag.Get(xlsTag)
		if len(tagStr) == 0 {
			continue
		}

		field := dataType.Field(i)
		fieldName := field.Name

		option := optionMap[fieldName]

		fieldValue := dataValue.FieldByName(fieldName)
		if !fieldValue.IsValid() {
			continue
		}

		switch field.Type.Kind() {
		case reflect.Ptr, reflect.Struct:
			if fieldValue.CanInterface() {
				if option != nil && option.Inline {
					err = addRow(row, fieldValue.Interface(), optionMap)
					if err != nil {
						panic(err)
					}
				}
			}

		case reflect.Interface, reflect.Map, reflect.Array, reflect.Slice, reflect.Complex64, reflect.Complex128, reflect.UnsafePointer, reflect.Chan, reflect.Func:
			cell := row.AddCell()
			cell.Value = "# unsupported by xlsx-util #"

		case reflect.Float64, reflect.Float32:
			cell := row.AddCell()
			cell.SetType(xlsx.CellTypeNumeric)
			precision := defaultFloatPrecision
			if option != nil && option.Precision > 0 {
				precision = option.Precision
			}

			if fieldValue.CanInterface() {
				if field.Type.Kind() == reflect.Float32 {
					cell.Value = strconv.FormatFloat(float64(fieldValue.Interface().(float32)), 'f', precision, 64)
				} else {
					cell.Value = strconv.FormatFloat(fieldValue.Interface().(float64), 'f', precision, 64)
				}
			}

		case reflect.Int8, reflect.Uint8, reflect.Int16, reflect.Uint16, reflect.Int32, reflect.Uint32,
			reflect.Int64, reflect.Uint64, reflect.Int, reflect.Uint:
			cell := row.AddCell()
			cell.SetType(xlsx.CellTypeNumeric)
			if fieldValue.CanInterface() {
				cell.Value = fmt.Sprintf("%v", fieldValue.Interface())
			}

		case reflect.String:
			cell := row.AddCell()

			if fieldValue.CanInterface() {
				cell.Value = fmt.Sprintf("%v", fieldValue.Interface())

				if len(cell.Value) == 8 && cell.Value[:2] == "20" { // 特殊处理日期 20060102 的情况，
					_, err = time.Parse("20060102", cell.Value)
					if err == nil {
						cell.SetType(xlsx.CellTypeNumeric)
					}
					_, err = time.Parse("2006/01/02", cell.Value)
					if err == nil {
						cell.SetType(xlsx.CellTypeDate)
					}

				} else {
					cell.SetType(xlsx.CellTypeString)
				}

			}

		}
	}
	return nil
}

func getOptionFromTag(tag string) *xlsOption {
	// construct a yaml str
	yamlStr := ""
	tagStrs := strings.Split(tag, ",")

	option := &xlsOption{}

	if len(tagStrs) <= 1 {
		return new(xlsOption)
	}

	for _, v := range tagStrs[1:] {
		tagStrSeg := v
		if inlineTag == strings.TrimSpace(tagStrSeg) {
			tagStrSeg += ":true"
		}
		segs := strings.Split(strings.TrimSpace(tagStrSeg), ":")
		yamlStr += strings.TrimSpace(segs[0]) + ": " + strings.TrimSpace(segs[1])
	}
	err := yaml.Unmarshal([]byte(yamlStr), option)
	if err != nil {
		panic(err)
	}

	return option
}

// just to construct a yaml document, and deserialize via yaml package
// data: pointer of a slice
func XlsBindByYamlTag(file *xlsx.File, sheetName string, data interface{}) error {
	sheet, ok := file.Sheet[sheetName]
	if !ok {
		return errors.New("sheetName not found")
	}

	headerMap := make(map[int]string)

	yamlStr := bytes.Buffer{}
	for rowIndex, row := range sheet.Rows {
		if rowIndex == 0 {
			for columnIndex, cell := range row.Cells {
				if len(cell.String()) == 0 { // if value is empty, ignore
					continue
				}
				headerMap[columnIndex] = cell.String()
			}
			continue
		}
		// construct a yaml list format document
		yamlStr.WriteString("- ")

		// check if this row is empty
		isRowEmpty := true
		for _, cell := range row.Cells {
			if len(cell.String()) > 0 {
				isRowEmpty = false
				break
			}
		}

		if isRowEmpty {
			continue
		}

		for columnIndex, cell := range row.Cells {
			if len(headerMap[columnIndex]) == 0 { // if head is empty, ignore
				continue
			}

			if columnIndex > 0 {
				yamlStr.WriteString("  ")
			}
			yamlStr.WriteString(headerMap[columnIndex] + `: ` + cell.String() + "\n")

		}
	}
	// you can open debug log here
	//fmt.Println(yamlStr.String())
	err := yaml.Unmarshal(yamlStr.Bytes(), data)
	return err
}
