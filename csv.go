package xlsxutil

import (
	"bufio"
	"bytes"
	"errors"
	"os"
	"reflect"
	"strings"
)

// dump data in csv format
// data: ptr of slice, slice element should be a struct
// sep: separator of csv line. be careful to avoid value conflict
func CsvDump(sep string, data interface{}) (*bytes.Buffer, error) {
	dataValue := reflect.ValueOf(data)
	if dataValue.Kind() != reflect.Slice {
		return nil, errors.New("do not support non slice data")
	}

	sliceLen := dataValue.Len()
	if sliceLen == 0 {
		return nil, errors.New("empty slice")
	}

	var optionMap map[string]*xlsOption
	buf := &bytes.Buffer{}
	var err error
	for i := 0; i < sliceLen; i++ {
		itemValue := dataValue.Index(i)
		if itemValue.Kind() == reflect.Ptr {
			itemValue = itemValue.Elem()
		}
		var rowStrSegs []string
		if i == 0 {
			optionMap, err = addHeaderRow(itemValue, func(str string, kind reflect.Kind) {
				rowStrSegs = append(rowStrSegs, str)
			})
			if err != nil {
				return nil, err
			}
			buf.WriteString(getLineFromRowSegs(rowStrSegs, sep) + "\n")
			continue
		}

		err = addRow(itemValue, optionMap, func(str string, kind reflect.Kind) {
			rowStrSegs = append(rowStrSegs, str)
		})
		if err != nil {
			return nil, err
		}
		buf.WriteString(getLineFromRowSegs(rowStrSegs, sep) + "\n")
	}
	return buf, nil
}

func getLineFromRowSegs(row []string, sep string) string {
	return strings.Join(row, sep)
}

// data: pointer of a slice
func CsvLoad(fileName string, sep string, data interface{}) error {
	file, err := os.Open(fileName)
	if err != nil {
		return err
	}
	defer file.Close()

	dataType, sliceValue, isElementPtr, err := validateDataInput(data)
	if err != nil {
		return err
	}

	dataValue := reflect.New(*dataType).Elem()
	_, optionMap := getStructOptions(dataValue)

	headerMap := make(map[int]string)

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

		rowStr := scanner.Text()
		if rowStr == "" {
			continue
		}
		valueMap := make(map[string]string)

		rowStrs := strings.Split(rowStr, sep)
		for k, v := range rowStrs {
			if len(headerMap[k]) == 0 { // if head is empty, ignore
				continue
			}

			valueMap[headerMap[k]] = v
		}

		addElement(*sliceValue, *dataType, isElementPtr, valueMap, optionMap)
	}

	return err
}
