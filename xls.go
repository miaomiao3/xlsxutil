package xlsxutil

import (
	"errors"
	"fmt"
	"github.com/miaomiao3/log"
	"github.com/miaomiao3/xlsx" // i fork this repo to enable SetType
	"gonum.org/v1/gonum/floats"
	"mtest/utils"
	"reflect"
	"strconv"
	"strings"
)

const (
	xlsTag                = "xls"
	inlineKey             = "inline"
	precisionKey          = "precision"
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
	XlsName   string
	IsInline  bool
	Precision int
}




type LineAppender interface {
	AddLine
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

		if field.Type.Kind() == reflect.Struct && option.IsInline {
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
				if option != nil && option.IsInline {
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
				cell.SetType(xlsx.CellTypeString)

			}

		}
	}
	return nil
}

func getStructOptions(dataValue reflect.Value) map[string]*xlsOption {
	if dataValue.Kind() == reflect.Ptr {
		dataValue = dataValue.Elem()
	}

	optionMap := make(map[string]*xlsOption)

	dataType := dataValue.Type()
	for i := 0; i < dataValue.NumField(); i++ {
		fieldValue := dataValue.Field(i)
		fieldType := dataType.Field(i)
		tag := fieldType.Tag.Get(xlsTag)

		optionMap[fieldType.Name] = getOptionFromTag(tag)
		switch fieldValue.Kind() {
		case reflect.Ptr:
			if fieldValue.Elem().Kind() == reflect.Struct {
				newOptionMap := getStructOptions(fieldValue.Elem())
				for k, v := range newOptionMap {
					optionMap[k] = v
				}
			}
		case reflect.Struct:
			newOptionMap := getStructOptions(fieldValue)
			for k, v := range newOptionMap {
				optionMap[k] = v
			}
		default:

		}
	}

	return optionMap
}

func getOptionFromTag(tag string) *xlsOption {
	tagStrs := strings.Split(tag, ",")

	option := &xlsOption{
		XlsName: tagStrs[0],
	}

	if len(tagStrs) <= 1 {
		return option
	}

	for _, v := range tagStrs[1:] {
		tagStrSeg := v
		if inlineKey == strings.TrimSpace(tagStrSeg) {
			option.IsInline = true
			continue
		}
		segs := strings.Split(strings.TrimSpace(tagStrSeg), ":")
		if precisionKey == strings.TrimSpace(segs[0]) {
			option.Precision, _ = strconv.Atoi(strings.TrimSpace(segs[1]))

		}
	}
	return option
}

var (
	errInputType = errors.New("wants pointer of slice, slice contains struct element")
)

// validate data, return slice element of struct, sliceValue, isElementPtr, error
func validateDataInput(data interface{}) (*reflect.Type, *reflect.Value, bool, error) {
	dataType := reflect.TypeOf(data)
	if dataType.Kind() != reflect.Ptr {
		return nil, nil, false, errInputType
	}

	dataType = dataType.Elem()

	if dataType.Kind() != reflect.Slice {
		return nil, nil, false, errInputType
	}

	sliceValue := reflect.ValueOf(data).Elem()
	dataType = dataType.Elem()
	isElementPtr := false
	if dataType.Kind() == reflect.Ptr {
		isElementPtr = true
		dataType = dataType.Elem()
	}

	if dataType.Kind() != reflect.Struct {
		return nil, nil, false, errInputType
	}

	return &dataType, &sliceValue, isElementPtr, nil

}

// load one sheet to slice
func XlsLoad(file *xlsx.File, sheetName string, data interface{}) error {
	sheet, ok := file.Sheet[sheetName]
	if !ok {
		return errors.New("sheetName not found")
	}

	dataType, sliceValue, isElementPtr, err := validateDataInput(data)
	if err != nil {
		return err
	}

	dataValue := reflect.New(*dataType).Elem()
	optionMap := getStructOptions(dataValue)

	log.Debug("optionMap", utils.GetJsonIdent(optionMap))

	// column index ->  column cell string
	headerMap := make(map[int]string)

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

		// check if this row is empty
		isRowEmpty := true
		for _, cell := range row.Cells {
			if len(strings.TrimSpace(cell.String())) > 0 {
				isRowEmpty = false
				break
			}
		}

		if isRowEmpty {
			break
		}

		valueMap := make(map[string]string)
		for columnIndex, cell := range row.Cells {
			if len(headerMap[columnIndex]) == 0 { // if head is empty, ignore
				continue
			}
			valueMap[headerMap[columnIndex]] = cell.String()
		}

		if isRowEmpty {
			break
		}

		log.Debug("valueMap", valueMap)
		addElement(*sliceValue, *dataType, isElementPtr, valueMap, optionMap)
	}

	return nil
}

func addElement(sliceValue reflect.Value, dataType reflect.Type, isPtr bool, valueMap map[string]string, optionMap map[string]*xlsOption) {
	var elem reflect.Value
	elem = reflect.New(dataType).Elem()

	setStructValue(elem, valueMap, optionMap)

	if isPtr {
		sliceValue.Set(reflect.Append(sliceValue, elem.Addr()))
	} else {
		sliceValue.Set(reflect.Append(sliceValue, elem))
	}

}

func setStructValue(dataValue reflect.Value, valueMap map[string]string, optionMap map[string]*xlsOption) {
	if dataValue.Kind() == reflect.Ptr {
		dataValue = dataValue.Elem()
	}

	if !dataValue.CanAddr() {
		return
	}
	dataType := dataValue.Type()
	for i := 0; i < dataValue.NumField(); i++ {
		fieldValue := dataValue.Field(i)
		fieldType := dataType.Field(i)

		option, ok := optionMap[fieldType.Name]
		if !ok {
			continue
		}

		fieldStr := valueMap[option.XlsName]
		if fieldStr == "" && !option.IsInline {
			continue
		}

		log.Debug("get tag  value ", fieldStr)
		switch fieldValue.Kind() {
		case reflect.Ptr:
			if fieldValue.Elem().Kind() == reflect.Struct {
				setStructValue(fieldValue.Elem(), valueMap, optionMap)
			}

		case reflect.Struct:
			setStructValue(fieldValue, valueMap, optionMap)

		case reflect.String:
			fieldValue.SetString(fieldStr)

		case reflect.Int,
			reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
			reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			numValue, err := strconv.ParseInt(fieldStr, 10, 64)
			if err != nil {
				log.Error("ParseInt of tag:%v err:%v", fieldValue, err)
				break
			}
			fieldValue.SetInt(numValue)

		case reflect.Float32, reflect.Float64:
			floatValue, err := strconv.ParseFloat(fieldStr, 64)
			if err != nil {
				log.Error("ParseInt of tag:%v err:%v", fieldValue, err)
				break
			}
			if option.Precision > 0 {
				floatValue = floats.Round(floatValue, option.Precision)
			}
			fieldValue.SetFloat(floatValue)

		default:
		}
	}
}
