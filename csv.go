package xlsxutil

import (
	"bufio"
	"bytes"
	"errors"
	"fmt"
	"gopkg.in/yaml.v2"
	"os"
	"reflect"
	"strconv"
	"strings"
)

// dump data in csv format
// data: ptr of slice, slice element should be a struct
// sep: separator of csv line. be careful to avoid value conflict
func CsvDump(sep string, data interface{}) (*bytes.Buffer, error) {
	value := reflect.ValueOf(data)
	buf := &bytes.Buffer{}
	if value.Kind() != reflect.Slice {
		return buf, errors.New("do not support non slice data")
	}
	l := value.Len()
	if l == 0 {
		return buf, errors.New("empty slice")
	}
	dataSlice := make([]interface{}, l)
	for i := 0; i < l; i++ {
		if value.Index(i).CanInterface() {

			if value.Index(i).IsNil() {
				return buf, errors.New("slice contain nil element")
			}

			//if value.Index(i).CanAddr() {
			//	return errors.New("slice contains element that can't addr")
			//}
			dataSlice[i] = value.Index(i).Interface()
		} else {
			return buf, errors.New("slice item CanInterface failed")
		}

	}

	var optionMap map[string]*xlsOption
	var rowStrSegs []string
	var err error
	for i := 0; i < l; i++ {
		if i == 0 {
			// get header row and its option
			optionMap, rowStrSegs, err = getCsvHeaderOption(dataSlice[i])
			if err != nil {
				return buf, err
			}
			buf.WriteString(getLineFromRowSegs(rowStrSegs, sep) + "\n")
		}
		rowStrSegs, err = getLineFromStruct(dataSlice[i], optionMap)
		if err != nil {
			return buf, err
		}
		buf.WriteString(getLineFromRowSegs(rowStrSegs, sep) + "\n")
	}
	return buf, nil
}

func getLineFromRowSegs(row []string, sep string) string {
	return strings.Join(row, sep)
}

// get field option using 'xls' tag
// generate header line
func getCsvHeaderOption(v interface{}) (optionMap map[string]*xlsOption, row []string, err error) {
	dataValue := reflect.ValueOf(v)
	dataType := reflect.TypeOf(v)
	if dataType == nil {
		return optionMap, nil, fmt.Errorf("getCsvHeaderOption get nil input")
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
			newOptionMap, embedRow, err := getCsvHeaderOption(fieldValue.Interface())
			if err != nil {
				return nil, nil, err
			}
			row = append(row, embedRow...)
			for k, v := range newOptionMap {
				optionMap[k] = v
			}
		} else {
			row = append(row, strings.Split(tag, ",")[0])
		}
	}
	return optionMap, row, nil
}

// change one struct to row value， 必须传结构体指针
func getLineFromStruct(v interface{}, optionMap map[string]*xlsOption) ([]string, error) {
	dataType := reflect.TypeOf(v)
	dataValue := reflect.ValueOf(v)
	var row []string

	if dataType.Kind() == reflect.Ptr {
		// only support struct ptr
		originType := reflect.ValueOf(v).Elem().Type()
		if originType.Kind() != reflect.Struct {
			err := errors.New("only support struct pointer")
			return nil, err
		}
		dataValue = dataValue.Elem()
		dataType = dataType.Elem()
	} else if dataType.Kind() != reflect.Struct {
		err := errors.New("only support struct")
		return row, err
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
					embedRow, err := getLineFromStruct(fieldValue.Interface(), optionMap)
					if err != nil {
						panic(err)
					}
					row = append(row, embedRow...)
				}
			}

		case reflect.Interface, reflect.Map, reflect.Array, reflect.Slice, reflect.Complex64, reflect.Complex128, reflect.UnsafePointer, reflect.Chan, reflect.Func:
			row = append(row, "# unsupported by xlsx-util #")

		case reflect.Float64, reflect.Float32:
			precision := defaultFloatPrecision
			if option != nil && option.Precision > 0 {
				precision = option.Precision
			}
			if fieldValue.CanInterface() {
				if field.Type.Kind() == reflect.Float32 {
					newColumnStr := strconv.FormatFloat(float64(fieldValue.Interface().(float32)), 'f', precision, 64)
					row = append(row, newColumnStr)
				} else {
					newColumnStr := strconv.FormatFloat(fieldValue.Interface().(float64), 'f', precision, 64)
					row = append(row, newColumnStr)
				}
			}

		case reflect.Int8, reflect.Uint8, reflect.Int16, reflect.Uint16, reflect.Int32, reflect.Uint32,
			reflect.Int64, reflect.Uint64, reflect.Int, reflect.Uint:
			if fieldValue.CanInterface() {
				row = append(row, fmt.Sprintf("%v", fieldValue.Interface()))
			}

		case reflect.String:
			if fieldValue.CanInterface() {
				row = append(row, fmt.Sprintf("%v", fieldValue.Interface()))
			}

		}
	}
	return row, nil
}

// just to construct a yaml document, and deserialize via yaml package
// data: pointer of a slice
func CsvBindByYamlTag(fileName string, sep string, data interface{}) error {
	headerMap := make(map[int]string)

	yamlStr := bytes.Buffer{}

	file, err := os.Open(fileName)
	if err != nil {
		return err
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	lineCnt := 0
	for scanner.Scan() {
		lineCnt++
		if lineCnt == 1 {
			rowStr := scanner.Text()
			rowStrs := strings.Split(rowStr, sep)
			for columnIndex, cell := range rowStrs {
				if len(cell) == 0 { // if value is empty, ignore
					continue
				}
				headerMap[columnIndex] = cell
			}
			continue
		}

		// construct a yaml list format document
		yamlStr.WriteString("- ")

		rowStr := scanner.Text()
		rowStrs := strings.Split(rowStr, sep)
		for columnIndex := 0; columnIndex < len(rowStrs); columnIndex++ {
			cell := rowStrs[columnIndex]
			if len(headerMap[columnIndex]) == 0 { // if head is empty, ignore
				continue
			}

			if columnIndex > 0 {
				yamlStr.WriteString("  ")
			}
			yamlStr.WriteString(headerMap[columnIndex] + `: ` + cell + "\n")
		}
	}

	// you can open debug log here
	//fmt.Println(yamlStr.String())
	err = yaml.Unmarshal(yamlStr.Bytes(), data)
	return err
}
