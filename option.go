package xlsxutil

import (
	"errors"
	"fmt"
	"gonum.org/v1/gonum/floats"
	"reflect"
	"strconv"
	"strings"
)

const (
	xlsTag                = "xls"
	inlineKey             = "inline"
	precisionKey          = "precision"
	defaultFloatPrecision = -1
)

type xlsOption struct {
	XlsName   string
	IsInline  bool
	Precision int
}

func getStructOptions(dataValue reflect.Value) (fieldNames []string, optionMap map[string]*xlsOption) {
	if dataValue.Kind() == reflect.Ptr {
		dataValue = dataValue.Elem()
	}

	if optionMap == nil {
		optionMap = make(map[string]*xlsOption)
	}

	dataType := dataValue.Type()
	for i := 0; i < dataValue.NumField(); i++ {
		fieldValue := dataValue.Field(i)
		fieldType := dataType.Field(i)
		tag := fieldType.Tag.Get(xlsTag)

		optionMap[fieldType.Name] = getOptionFromTag(tag)
		switch fieldValue.Kind() {
		case reflect.Ptr:
			if fieldValue.Elem().Kind() == reflect.Struct {
				newFieldNames, newOptionMap := getStructOptions(fieldValue.Elem())
				for k, v := range newOptionMap {
					optionMap[k] = v
				}
				fieldNames = append(fieldNames, newFieldNames...)
			}
		case reflect.Struct:
			newFieldNames, newOptionMap := getStructOptions(fieldValue)
			for k, v := range newOptionMap {
				optionMap[k] = v
			}
			fieldNames = append(fieldNames, newFieldNames...)
		default:
			fieldNames = append(fieldNames, optionMap[fieldType.Name].XlsName)
		}
	}

	return
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

type rowHandleFunc func(str string, kind reflect.Kind)

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
				break
			}
			fieldValue.SetInt(numValue)

		case reflect.Float32, reflect.Float64:
			floatValue, err := strconv.ParseFloat(fieldStr, 64)
			if err != nil {
				break
			}
			if option.Precision > 0 {
				floatValue = floats.Round(floatValue, option.Precision)
			}
			fieldValue.SetFloat(floatValue)

		}
	}
}

// add header
func addHeaderRow(dataValue reflect.Value, f rowHandleFunc) (optionMap map[string]*xlsOption, err error) {
	var fieldNames []string
	fieldNames, optionMap = getStructOptions(dataValue)
	for _, v := range fieldNames {
		f(v, reflect.String)
	}

	return optionMap, nil
}

func addRow(dataValue reflect.Value, optionMap map[string]*xlsOption, f rowHandleFunc) (err error) {
	if dataValue.Kind() == reflect.Ptr {
		dataValue = dataValue.Elem()
	}

	if dataValue.Kind() != reflect.Struct {
		err = errors.New("only support struct or struct pointer")
		return
	}
	dataType := dataValue.Type()

	num := dataValue.NumField()
	for i := 0; i < num; i++ {
		fieldValue := dataValue.Field(i)
		fieldName := dataType.Field(i).Name

		option, ok := optionMap[fieldName]
		if !ok {
			continue
		}

		switch fieldValue.Kind() {
		case reflect.Ptr:
			if fieldValue.Elem().Kind() == reflect.Struct {
				if option.IsInline {
					err = addRow(fieldValue.Elem(), optionMap, f)
					if err != nil {
						return err
					}
				}
			}

		case reflect.Struct:
			if fieldValue.CanInterface() {
				if option.IsInline {
					err = addRow(fieldValue, optionMap, f)
					if err != nil {
						return err
					}
				}
			}

		case reflect.Interface, reflect.Map, reflect.Array, reflect.Slice, reflect.Complex64, reflect.Complex128, reflect.UnsafePointer, reflect.Chan, reflect.Func:
			f("# unsupported by xlsx-util #", fieldValue.Kind())

		case reflect.Float64, reflect.Float32:
			precision := defaultFloatPrecision
			if option != nil && option.Precision > 0 {
				precision = option.Precision
			}
			fieldStr := ""
			if fieldValue.CanInterface() {
				if fieldValue.Kind() == reflect.Float32 {
					fieldStr = strconv.FormatFloat(float64(fieldValue.Interface().(float32)), 'f', precision, 64)
				} else {
					fieldStr = strconv.FormatFloat(fieldValue.Interface().(float64), 'f', precision, 64)
				}
			}
			f(fieldStr, fieldValue.Kind())

		case reflect.Int8, reflect.Uint8, reflect.Int16, reflect.Uint16, reflect.Int32, reflect.Uint32,
			reflect.Int64, reflect.Uint64, reflect.Int, reflect.Uint, reflect.String:
			if fieldValue.CanInterface() {
				fieldStr := fmt.Sprintf("%v", fieldValue.Interface())
				f(fieldStr, fieldValue.Kind())
			}

		}
	}
	return nil
}
